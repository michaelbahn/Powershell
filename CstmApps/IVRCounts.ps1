#
#
#
#
#
#

$ftpfile =        "\\dgvmappuipd01\iFile-iBatch\DE4581\IVR\FTPL\"
$nonftpfile = "\\dgvmappuipd01\iFile-iBatch\DE4581\IVR\NON-FTPL"

$fileName =  "\\dgvmopspd01\DMSOPSWebSite\CstmAppsSrvcs.htm"
#"\\dgvmopspd01\wwwroot\CstmAppsSrvcs.htm"
#"\\dgvmopspd01\wwwroot\CCR_HTML.htm"
#New-Item -itemtype file -Path $fileName -Force
$date = Get-Date

Function writeHtmlHeader  
{  
param($fileName)  
$date = Get-Date  
Add-Content $fileName "<html>" 
Add-Content $fileName "<head>" 
Add-Content $fileName "<META HTTP-EQUIV=refresh CONTENT=15>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $fileName '<title>UI IVR FTP Monitor - DMS</title>' 
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
#Add-Content $FileName  "<a href='http://dgvmopspd01/default.aspx'>Back</a></td>"
add-content $fileName  "<td colspan='7' height='25' align='center'>" 
add-content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>Custom Application File Counts</strong></font>   $date" 
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

writeTableHeader $fileName  

  
$FTPCount = (Get-ChildItem -Path $ftpfile).count
if($ftpcount -le 0)
{
	Write-Output "FTP No Files. Count is " $FTPCount
	
	Add-Content $fileName "<tr bgcolor=#FF0000>" 
 		Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'>$ftpfile</td>"
 		Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'>$ftpcount</td>" 
 		Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'><a href=file:$ftpfile>Directory</a></td>"
		#$oldftpcount</td>" 																								<a href=file:'$ftpfile'>Directory</a></td>
 		Add-Content $fileName "</tr>" 
		
}
elseif($ftpcount -gt 0)
{
	Write-Output "FTP More than zero count" $ftpcount 
	
	Add-Content $fileName "<tr bgcolor='#D0E6FF'>" 
 		Add-Content $fileName "<td width='25%' align='center'>$ftpfile</td>" 
 		Add-Content $fileName "<td width='25%' align='center'>$ftpcount</td>" 
 		Add-Content $fileName "<td width='25%' align='center'><a href='file:$ftpfile'>Directory</a></td>" 
 		Add-Content $fileName "</tr>" 
	
}

$NonFTPCount = (Get-ChildItem -Path $nonftpfile).count
if($Nonftpcount -le 0)
{
	Write-Output "NON_FTP No Files. Count is " $NonFTPCount
	Add-Content $fileName "<tr bgcolor=#FF0000>" 
 		Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'>$nonftpfile</td>" 
 		Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'>$nonftpcount</td>" 
 		Add-Content $fileName "<td width='25%' align='center'><font color='#FFFF00'><a href=file:'$nonftpfile'>Directory</a></td>"
		Add-Content $fileName "</tr>" 
}

elseif($Nonftpcount -gt 0)
{
	Write-Output "NON-FTP More than zero count" $Nonftpcount 
	$oldnonftpcount = $nonFTPCount
		Add-Content $fileName "<tr bgcolor='#D0E6FF'>" 
 		Add-Content $fileName "<td width='25%' align='center'>$nonftpfile</td>" 
 		Add-Content $fileName "<td width='25%' align='center'>$nonftpcount</td>" 
 		Add-Content $fileName "<td width='25%' align='center'><a href=file:$nonftpfile>Directory</a></td>"
 		Add-Content $fileName "</tr>" 
}

exit
