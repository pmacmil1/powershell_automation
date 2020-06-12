#You will need the Sharepoint Client CSOM Assembilies to run this code
#The Windows installer can be found here: https://www.microsoft.com/en-us/download/details.aspx?id=42038

Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"
$Cred= Get-Credential

Function Update-UserProfileProperty()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $AdminCenterURL,
        [Parameter(Mandatory=$true)] [string] $UserAccount,
        [Parameter(Mandatory=$true)] [string] $PropertyName,
        [Parameter(Mandatory=$true)] [string] $PropertyValue
    )    
    Try {
        #Setup Credentials to connect
        
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
 
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($AdminCenterURL)
        $Ctx.Credentials = $Credentials
         
        #Get the User
        $User = $Ctx.web.EnsureUser($UserAccount)
        $Ctx.Load($User)
        $Ctx.ExecuteQuery()
 
        #Get User Profile
        $PeopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($Ctx)
         
        #update User Profile Property
        $PeopleManager.SetSingleValueProfileProperty($User.LoginName, $PropertyName, $PropertyValue)
        $Ctx.ExecuteQuery()
 
        #Write-host "User Profile Property has been Updated!" -f Green
    }
    Catch {
        write-host -f Red "Error Updating User Profile Property!" $_.Exception.Message
    }
}

#Define Parameter values
#NOTE: YOU HAVE TO USE THE ADMIN PAGE HERE, NOT THE REGULAR SHAREPOINT ONLINE URL FOR YOUR TENANT!
$AdminCenterURL="https://TENANT_NAME-admin.sharepoint.com"
$SipAddress = 'SPS-SipAddress'
$WorkPhone = 'WorkPhone'
$MobilePhone = 'CellPhone'

$Users = Import-CSV -Delimiter ';' -Path "C:\PATH_TO_CSV\SOME_CSV.CSV" -Encoding utf8

ForEach($User in $Users)
{
    #'Updateing User: '+$User.UserPrincipalName+' with extension number: '+$User.extensionNumber
    $SPUserName = 'i:0#.f|membership|'+$User.UserPrincipalName
    
    if($user.extensionNumber -eq "")
    {
        'Updateing User: '+ $SPUserName + ' with the SipAddress of ' + '-'
        Update-UserProfileProperty -AdminCenterURL $AdminCenterURL -UserAccount $SPUserName -PropertyName $SipAddress -PropertyValue '-'
    }
    else
    {
        'Updateing User: '+ $SPUserName + ' with the SipAddress of ' + $user.extensionNumber
        Update-UserProfileProperty -AdminCenterURL $AdminCenterURL -UserAccount $SPUserName -PropertyName $SipAddress -PropertyValue $User.extensionNumber
    }

    if($user.NewPhoneNumber -eq "")
    {
        'Updateing User: '+ $SPUserName + ' with the NewPhoneNumber of ' + '-'
        Update-UserProfileProperty -AdminCenterURL $AdminCenterURL -UserAccount $SPUserName -PropertyName $WorkPhone -PropertyValue '-'
    }
    else
    {
        'Updateing User: '+ $SPUserName + ' with the NewPhoneNumber of ' + $user.NewPhoneNumber
        Update-UserProfileProperty -AdminCenterURL $AdminCenterURL -UserAccount $SPUserName -PropertyName $WorkPhone -PropertyValue $User.NewPhoneNumber
    }

    if($user.MobilePhone -eq "")
    {
        'Updateing User: '+ $SPUserName + ' with the MobilePhone of ' + '-'
        Update-UserProfileProperty -AdminCenterURL $AdminCenterURL -UserAccount $SPUserName -PropertyName $MobilePhone -PropertyValue '-'
    }
    else
    {
        'Updateing User: '+ $SPUserName + ' with the MobilePhone of ' + $user.MobilePhone
        Update-UserProfileProperty -AdminCenterURL $AdminCenterURL -UserAccount $SPUserName -PropertyName $MobilePhone -PropertyValue $User.MobilePhone
    }
}
