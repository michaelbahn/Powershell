        try
        {
            regedit.exe /s "\\Dgvmopspd02\Deploy\iCapture\impression.reg" 
            regedit.exe /s "\\Dgvmopspd02\Deploy\iCapture\odbc-preprod.reg"
         }
         catch
        {
            return $_
        }
