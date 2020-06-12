#You will need the Sharepoint Client CSOM Assembilies to run this code
#The Windows installer can be found here: https://www.microsoft.com/en-us/download/details.aspx?id=42038

Add-Type -Path "C:\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

Function Ensure-Folder()
{
Param(
  [Parameter(Mandatory=$True)]
  [Microsoft.SharePoint.Client.Web]$Web,

  [Parameter(Mandatory=$True)]
  [Microsoft.SharePoint.Client.Folder]$ParentFolder, 

  [Parameter(Mandatory=$True)]
  [String]$FolderUrl

)
    $folderNames = $FolderUrl.Trim().Split("/",[System.StringSplitOptions]::RemoveEmptyEntries)
    $folderName = $folderNames[0]
    Write-Host "Creating folder [$folderName] ..."
    $curFolder = $ParentFolder.Folders.Add($folderName)
    $Web.Context.Load($curFolder)
    $web.Context.ExecuteQuery()
    Write-Host "Folder [$folderName] has been created succesfully. Url: $($curFolder.ServerRelativeUrl)"

    if ($folderNames.Length -gt 1)
    {
        $curFolderUrl = [System.String]::Join("/", $folderNames, 1, $folderNames.Length - 1)
        Ensure-Folder -Web $Web -ParentFolder $curFolder -FolderUrl $curFolderUrl
    }
}


Function Upload-File() 
{
Param(
  [Parameter(Mandatory=$True)]
  [Microsoft.SharePoint.Client.Web]$Web,

  [Parameter(Mandatory=$True)]
  [String]$FolderRelativeUrl, 

  [Parameter(Mandatory=$True)]
  [System.IO.FileInfo]$LocalFile

)

    try 
    {
       $FileExtension = [System.IO.Path]::GetExtension($LocalFile.Name)
	   #Confirm the file type, that you want to upload here
	   #Currently set for PDFs
       if ($FileExtension -eq ".pdf")
       {
           
           $fileUrl = $FolderRelativeUrl + "/" + $LocalFile.Name
           Write-Host "Uploading file [$($LocalFile.FullName)] ... to $fileUrl"
           [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($Web.Context, $fileUrl, $LocalFile.OpenRead(), $true)
           Write-Host "File [$($LocalFile.FullName)] has been uploaded succesfully. Url: $fileUrl"
       }
       else
       {
         write-host "This file is not a pdf [$($LocalFile.FullName)]"
       }
    }
    catch 
    {
       write-host "An error occured while uploading file [$($LocalFile.FullName)]"
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $ErrorMessage
        $FailedItem
    }
}


function Upload-Files()
{

Param(
  [Parameter(Mandatory=$True)]
  [String]$Url,

  [Parameter(Mandatory=$True)]
  [String]$UserName,

  [Parameter(Mandatory=$False)]
  [String]$Password, 

  [Parameter(Mandatory=$True)]
  [String]$TargetListTitle,

  [Parameter(Mandatory=$True)]
  [String]$SourceFolderPath

)

    if($Password) {
       $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    }
    else {
      $SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString
    }
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,$SecurePassword)
    $Context.Credentials = $Credentials
    $web = $Context.Web 
    $Context.Load($web)
    $list = $web.Lists.GetByTitle($TargetListTitle);
    $Context.Load($list.RootFolder)
    $Context.ExecuteQuery()


    $Date = (get-date).ToString(‘dd-MM-yy’)

    
    $curFolder = $list.RootFolder.Folders.Add($Date)
    $Web.Context.Load($curFolder)
    $web.Context.ExecuteQuery()

    $folderRelativeUrl = $folderRelativeUrl + "/"+ $Date
    $DateFolder = $curFolder
    $DateFolder = [Microsoft.SharePoint.Client.Folder]$DateFolder

    Get-ChildItem $SourceFolderPath -Recurse | % {
       if ($_.PSIsContainer -eq $True) {
          $folderUrl = $_.FullName.Replace($SourceFolderPath,"").Replace("\","/")   
          if($folderUrl) {
             Ensure-Folder -Web $web -ParentFolder $DateFolder -FolderUrl $folderUrl
          }
       }
       else{
          $Date = (get-date).ToString(‘dd-MM-yy’)
          $folderRelativeUrl = $list.RootFolder.ServerRelativeUrl + "/" + $Date + $_.DirectoryName.Replace($SourceFolderPath,"").Replace("\","/")
          $folderRelativeUrl
          Upload-File -Web $web -FolderRelativeUrl $folderRelativeUrl -LocalFile $_ 
       }
    }
}

#Usage
$LogfileDate = (get-date).ToString(‘d-M-y’)
#Change the directory for the log file
$Logfile = "\\FILESHARE\Sharepoint Uploader\SharepointUploader - "+$LogfileDate+".log"
Start-Transcript $LogFile

$Url = "https://DOMAIN_NAME.sharepoint.com/sites/SITE_NAME/"

$TargetListTitle = "TARGET"  #Target Library


$SourceFolderPath = "C:\FILES_TO_UPLOAD"  #Source Physical Path 

#Upload files
Upload-Files -Url $Url -UserName $UserName -Password $Password -TargetListTitle $TargetListTitle -SourceFolderPath $SourceFolderPath

Stop-Transcript
