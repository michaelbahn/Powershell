$getdate = get-date -format "yyyy-MM-dd"
$date = "$getdate 00:00:00"
$iCapture = (Invoke-SQLcmd -ServerInstance "DGVMSQLPD08\DGC1VDBRPTPD01" -Query "Select Sum(BRD.schBRD.tblReconciliation.iCapture_Docs) as icap_data FROM [BRD].[schBRD].[tblReconciliation] where Received_Date = '$date'").icap_data
$InfoImage = (Invoke-SQLcmd -ServerInstance "DGVMSQLPD08\DGC1VDBRPTPD01" -Query "Select Sum(BRD.schBRD.tblReconciliation.InfoImage_Docs) as infoI_data FROM [BRD].[schBRD].[tblReconciliation] where Received_Date = '$date'").infoi_data
If ($iCapture -eq $InfoImage) { 
$report = "Nothing to report!    iCapture count = $iCapture   and   InfoImage count = $InfoImage match"
Send-MailMessage -From 'TeamOps <teamops@edd.ca.gov>' -To 'TeamOps <teamops@edd.ca.gov>' -Subject "BRD Report $getdate - Nothing to Report!" -body $report -SmtpServer 'smtp.edd.ca.gov'
} else { 
If ($iCapture -ne $InfoImage) {
$report = "ATTN needed!    iCapture count = $iCapture   and   InfoImage count = $InfoImage do NOT match"
Send-MailMessage -From 'TeamOps <teamops@edd.ca.gov>' -To 'TeamOps <teamops@edd.ca.gov>' -Subject "BRD Report $getdate - ATTN Needed!" -body $report -SmtpServer 'smtp.edd.ca.gov'
}
}
