cls

$rbeDir = "\\dgvmrbebendv01\d$\RBE01"

$files = ("error.sts", "docssent.sts", "finished.sts", "verify.sts")

foreach ($file in $files)
{
    Write-Host $file
    #clean up directory
    $filesToDelete = Get-ChildItem -Path $rbeDir -File $file -Recurse 
    Write-Host $filesToDelete.Count
    $filesToDelete | Remove-Item -Force -Recurse

}



