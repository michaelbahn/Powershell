function Query-Transmission-Status ($sqlServer)
{
    $query = "use FSET;
    	SELECT	dbo.vTransmissionDetail.TransmissionID,
			dbo.vTransmissionDetail.TransmissionConfirmationNumber, 
			dbo.vTransmissionDetail.PostMarkedDate, 
			DateDiff (minute, dbo.vTransmissionDetail.PostMarkedDate, 
			dbo.vTransmissionDetail.ModifiedDate) AS ProcessingTime,
			dbo.vTransmissionStatus.TransmissionStatusDescription
	FROM	dbo.vTransmission
			INNER JOIN dbo.vTransmissionDetail 
			ON dbo.vTransmissionDetail.TransmissionID = dbo.vTransmission.TransmissionID 
			INNER JOIN [dbo].[vTransmissionStatus]
			ON dbo.vTransmissionDetail.TransmissionStatusCode = [dbo].[vTransmissionStatus].TransmissionStatusCode		
	WHERE   dbo.vTransmissionDetail.TransmissionStatusCode IN ('TRPS', 'TRMF')
		 AND   dbo.vTransmissionDetail.PostMarkedDate > DATEADD(day, -180, GETDATE())"

    return Invoke-Sqlcmd -ServerInstance $sqlServer -Query $query
    
}

function Query-Transmission-Elements  ($sqlServer, $transmissionID)
{
    $query = "use FSET;
    SELECT dbo.vTransmissionElement.FormTypeCode, dbo.vTransmissionElement.TransmissionID, 
           dbo.vTransmissionElement.ConfirmationNumber, dbo.vTransmissionElement.ModifiedDate, 
           dbo.vTransmissionElementDetail.BatchNumber, dbo.vTransmissionElementDetail.ErrorCode, dbo.vTransmissionDetail.ProcessID, 
           dbo.vTransmissionElementDetail.TransmissionElementStatusCode as StatusCode, dbo.vTransmissionElementDetail.ErrorFilePath
    FROM   dbo.vTransmissionElement 
		    INNER JOIN dbo.vTransmissionElementDetail ON dbo.vTransmissionElement.TransmissionElementID = dbo.vTransmissionElementDetail.TransmissionElementID 
		    INNER JOIN dbo.vTransmission ON dbo.vTransmissionElement.TransmissionID = dbo.vTransmission.TransmissionID 
		    INNER JOIN dbo.vTransmissionDetail ON dbo.vTransmission.TransmissionID = dbo.vTransmissionDetail.TransmissionID
    WHERE  dbo.vTransmissionElement.TransmissionID  = $($transmissionID)"

    return Invoke-Sqlcmd -ServerInstance $sqlServer -Query $query
    
}

