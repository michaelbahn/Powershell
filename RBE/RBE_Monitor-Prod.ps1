$dir = $MyInvocation.MyCommand.Path
$scriptPath  = Split-Path $dir
Set-Location  $scriptPath

# First lets create a text file, where we will later save the DIA HTML status  
Import-Module ..\Logging.psm1

$VerbosePreference = 'Continue'
$DebugPreference = 'Continue'
$LogFilePreference = "\\dgvmopspd01\logs\RBEMon.log"

Write-HostLog "Host output"
Write-VerboseLog "Verbose output"
Write-DebugLog "Debug output"
Write-WarningLog "Warning output"
$ErrorActionPreference = 'silentlycontinue'

$rbes = Import-Csv "\\dgvsopspd0`\scripts\RBE\RBE.csv"
$FileName = "\\dgvmopspd01\wwwroot\RBEMon.htm" 
#$Logfilename = "\\DGVSOPSPD01\Logs\RBEMon.txt"
$n = 1 

New-Item -ItemType file $FileName -Force 

$date = Get-Date | Write-HostLog 

# Function to write the HTML Header to the file  
# Add parameters for message and data needed to email
Function SendMail 
 	{
 	param($from, $to, $Subject, $Body)
	$date = Get-Date
	# Create from/to addresses  
	$from = New-Object System.Net.Mail.MailAddress "RBE@edd.ca.gov"  
	$to =   New-Object System.Net.Mail.MailAddress "teaminf@edd.ca.gov"
#	"grant.fraser@edd.ca.gov"
#	"teaminf@edd.ca.gov"
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
	$message.subject | Write-Hostlog 
	$message.body | Write-Hostlog 
	}
Function writeHtmlHeader  
{  
param($fileName)  
$date = Get-Date  
Add-Content $fileName "<html>" 
Add-Content $fileName "<head>" 
Add-Content $fileName "<META HTTP-EQUIV=refresh CONTENT=15>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $fileName '<title>RBE Monitor - All InfoImage Domains - Goethe</title>' 
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
Add-Content $FileName  "<a href='http://dgvmopspd01/default.aspx'>Back</a></td>"
add-content $fileName  "<td colspan='7' height='33' align='center'>" 
add-content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>RBE Monitor - All InfoImage Domains - Goethe - $date</strong></font>" 
add-content $fileName  "</td>" 
add-content $fileName  "</tr>" 
add-content $fileName  "</table>" 
}  

	
 # Function to write the HTML Header to the file  
 Function writeTableHeader  
 {  
 param($fileName)  
 add-content $fileName  "<table width='100%'>" 
 add-content $fileName  "<tr bgcolor='#D0E6FF'>" 
 Add-Content $fileName "<td width='20%' align='center'><strong>Server Name</strong></td>" 
 Add-Content $fileName "<td width='20%' align='center'><strong>Queue</strong></td>" 
 Add-Content $fileName "<td width='20%' align='center'><strong>Batches Queued</strong></td>"
 Add-Content $fileName "<td width='20%' align='center'><strong>Subdirectory</strong></td>"
 Add-Content $fileName "<td width='20%' align='center'><strong>RBEACK Status</strong></td>" 
 add-content $fileName  "</td>" 
 add-content $fileName  "</tr>" 
 add-content $fileName  "</table>"    
 }  
    
 Function writeHtmlFooter  
 {  
 param($fileName)  
# Add-Content $fileName "</body>" 
# Add-Content $fileName "</html>" 
 Add-Content $FileName "<table width='100%'><tbody>" 
 Add-Content $FileName "<tr bgcolor='#D0E6FF'>" 
 Add-Content $FileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>End Report</strong></font></td>" 
 Add-Content $FileName "</tr>"    
 Add-Content $fileName "</body>" 
 Add-Content $fileName "</html>" 
 }  


  
  writeHtmlHeader $FileName 


# This is the key part
##############################################################
	
	$date = Get-Date -Format g
	Write-Hostlog "------------------------------------------------------------------------"  
	Write-HostLog $date
	writetableheader $FileName
	
	
	foreach($rbe in $rbes)
		{
		sleep 1
		$server = $rbe.Server
		$Domain = $rbe.Domain
		$subdir = $rbe.SubDir
		$rbenum = $rbe.RBE
		$Batches = 0
#		Table Sub-Header
		Add-Content $FileName "<table width='100%'><tbody>" 
 		Add-Content $FileName "<tr bgcolor='#D0E6FF'>" 
  		Add-Content $FileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> $Domain </strong></font></td>" 
  		Add-Content $FileName "</tr>" 
		
		########################## Check to see if RBEACK is running ################################	
			$check=Get-Process -ComputerName $server  -Name "*RBEACK*" -ErrorAction SilentlyContinue
			if ($check -eq $null) 
			{
				$app = $check
				$Rbeack = "STOPPED"
				Write-HostLog "$app on $server is $RBEack"
#				Add-Content $fileName "<tr bgcolor=#FF0000>" 
#				Add-Content $fileName "<td width='20%' align='center'><strong><font color='#FFFF00'>$RBEack</td>"
# 				Add-Content $fileName "</tr>" 
			}
			else 
			{
				$Rbeack = "Running"
				Write-Hostlog "RBEACK on $server is $RBEack"
#				Add-Content $fileName "<td width='20%' align='center'><strong><font color='#003399'>$RBEack</td>" 
# 				Add-Content $fileName "</tr>"  
			}
			$path = $subdir + "\"
#			$path = $subdir + "\" + $rbenum
		
		$Batches = (Get-ChildItem $path -Filter "*.").count
				
		Write-Hostlog "$server in $Domain, $rbenum has $Batches Batches queued."
		
	If(($Batches -lt 40) -and ($RBEack -eq "Running"))
		{
		Add-Content $fileName "<tr bgcolor=#D0E6FF>" 
		Add-Content $fileName "<td width='20%' align='center'><font color='#003399'>$server</td>"
		Add-Content $fileName "<td width='20%' align='center'><font color='#003399'>$rbenum</td>" 
 		Add-Content $fileName "<td width='20%' align='center'><font color='#003399'><a href='file:$path'>Directory</a></td>" 
		Add-Content $fileName "<td width='20%' align='center'><font color='#003399'>$Batches</td>" 
		Add-Content $fileName "<td width='20%' align='center'><font color='#003399'>$RBEack</td>" 
 		Add-Content $fileName "</tr>" 
		}
#	
		elseif(($Batches -ge 41) -and ($batches -lt 70)) 
		{
		Add-Content $fileName "<tr bgcolor=#FFFF00>" 
 		Add-Content $fileName "<td width='20%' align='center'><strong><font color='003399'>$server</td>" 
 		Add-Content $fileName "<td width='20%' align='center'><strong><font color='#003399'>$rbenum</td>" 
		Add-Content $fileName "<td width='20%' align='center'><font color='#003399'><a href='file:$path'>Directory</a></td>" 
 		Add-Content $fileName "<td width='20%' align='center'><strong><font color='#003399'>$Batches</td>" 
		Add-Content $fileName "<td width='20%' align='center'><strong><font color='#003399'>$RBEack</td>" 
 		Add-Content $fileName "</tr>"  
 		}
		elseif (($Batches -ge 71) -or ($RBEack -eq "STOPPED"))
		{
		Add-Content $fileName "<tr bgcolor=#FF0000>" 
 		Add-Content $fileName "<td width='20%' align='center'><strong><font color='#FFFF00'>$server</td>" 
 		Add-Content $fileName "<td width='20%' align='center'><strong><font color='#FFFF00'>$rbenum</td>" 
		Add-Content $fileName "<td width='20%' align='center'><font color='#003399'><a href='file:$path'>Directory</a></td>" 
 		Add-Content $fileName "<td width='20%' align='center'><strong><font color='#FFFF00'>$Batches</td>" 
		Add-Content $fileName "<td width='20%' align='center'><strong><font color='#FFFF00'>$RBEack</td>"
 		Add-Content $fileName "</tr>" 
		$from = "RBE@EDD.ca.gov"
		$to = "Grant.fraser@edd.ca.gov"
		$Subject = "$domain $server $rbenum has a problem."
		$Body = "$batches Batches found; RBEACK is $rbeack."
		Sendmail $from $to $Subject $Body
		Write-HostLog $Body
		}
				
		
	Add-Content $fileName "<tr bgcolor=#D0E6FF>" 
 	Add-Content $fileName "<td width='20%' align='center'></td>" 
 	Add-Content $fileName "<td width='20%' align='center'></td>" 
 	Add-Content $fileName "<td width='20%' align='center'></td>"
	Add-Content $fileName "<td width='20%' align='center'></td>" 
	Add-Content $fileName "<td width='20%' align='center'></td>" 
 	Add-Content $fileName "</tr>" 
	$i++
	

	}


######sleep 15
writeHtmlFooter $FileName 

Invoke-Item -Path $FileName