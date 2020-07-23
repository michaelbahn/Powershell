#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
 
Function Copy-Files
{
  param
    (
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.Folder] $SourceFolder,
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.Folder] $TargetFolder
    )
    Try {
        #Get all Files from the source folder
        $SourceFilesColl = $SourceFolder.Files
        $SourceFolder.Context.Load($SourceFilesColl)
        $SourceFolder.Context.ExecuteQuery()
 
        #Iterate through each file and copy
        Foreach($SourceFile in $SourceFilesColl)
        {
            #Get the source file
            $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($SourceFolder.Context, $SourceFile.ServerRelativeUrl)
             
            #Copy File to the Target location
            $TargetFileURL = $TargetFolder.ServerRelativeUrl+"/"+$SourceFile.Name
            [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($TargetFolder.Context, $TargetFileURL, $FileInfo.Stream,$True)
 
            Write-host -f Green "Copied File '$($SourceFile.ServerRelativeUrl)' to '$TargetFileURL'"
        }
 
        #Process Sub Folders
        $SubFolders = $SourceFolder.Folders
        $SourceFolder.Context.Load($SubFolders)
        $SourceFolder.Context.ExecuteQuery()
        Foreach($SubFolder in $SubFolders)
        {
            If($SubFolder.Name -ne "Forms")
            {
                #Prepare Target Folder
                $TargetFolderURL = $SubFolder.ServerRelativeUrl -replace $SourceLibrary.RootFolder.ServerRelativeUrl, $TargetLibrary.RootFolder.ServerRelativeUrl
                Try {
                        $Folder=$TargetFolder.Context.web.GetFolderByServerRelativeUrl($TargetFolderURL)
                        $TargetFolder.Context.load($Folder)
                        $TargetFolder.Context.ExecuteQuery()
                    }
                catch {
                        #Create Folder
                        if(!$Folder.Exists)
                        {
                            $TargetFolderURL
                            $Folder=$TargetFolder.Context.web.Folders.Add($TargetFolderURL)
                            $TargetFolder.Context.Load($Folder)
                            $TargetFolder.Context.ExecuteQuery()
                            Write-host "Folder Added:"$SubFolder.Name -f Yellow
                        }
                    }
                #Call the function recursively
                Copy-Files -SourceFolder $SubFolder -TargetFolder $Folder
            }
        }
    }
    Catch {
        write-host -f Red "Error Copying File!" $_.Exception.Message
    }
}
 
#Set Parameter values
$SourceSiteURL="https://eddnet/teams/ASD/DMS_IT"
$TargetSiteURL="https://eddnet/teams/ISD/DTS/DMSOps/"
 
$SourceLibraryName="CSO Photos"
$TargetLibraryName="DMSO Photos"
 
#Setup Credentials to connect
$Cred= Get-Credential
$Credentials = New-Object System.Net.NetworkCredential($Cred.Username, $Cred.Password)
 
#Setup the contexts
$SourceCtx = New-Object Microsoft.SharePoint.Client.ClientContext($SourceSiteURL)
$SourceCtx.Credentials = $Credentials
$TargetCtx = New-Object Microsoft.SharePoint.Client.ClientContext($TargetSiteURL)
$TargetCtx.Credentials = $Credentials
      
#Get the source library and Target Libraries
$SourceLibrary = $SourceCtx.Web.Lists.GetByTitle($SourceLibraryName)
$SourceCtx.Load($SourceLibrary)
$SourceCtx.Load($SourceLibrary.RootFolder)
 
$TargetLibrary = $TargetCtx.Web.Lists.GetByTitle($TargetLibraryName)
$TargetCtx.Load($TargetLibrary)
$TargetCtx.Load($TargetLibrary.RootFolder)
$TargetCtx.ExecuteQuery()
 
#Call the function
Copy-Files -SourceFolder $SourceLibrary.RootFolder -TargetFolder $TargetLibrary.RootFolder


#Read more: http://www.sharepointdiary.com/2017/02/sharepoint-online-copy-files-between-site-collections-using-powershell.html#ixzz5uj1e0d4H