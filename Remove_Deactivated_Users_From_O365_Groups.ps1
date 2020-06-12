#############################################################
### Remove Azure AD users from all Office365/Azure Groups ###
#############################################################
function removeAzureADUserFromAllOffice365Groups
{
    Write-Host "===========================================================================" -ForegroundColor Yellow
    Write-Host "Removing deactivated Azure AD Users from all Office365 Groups and Mailboxes" -ForegroundColor Yellow
    Write-Host "===========================================================================" -ForegroundColor Yellow
    
    #Connect to Azure AD using the Connect-AzureAD method described above
    #Connect to Office 365 via AAD Graph
    #You may need to install the module as show below
    #Install-Module -Name AzureAD

    #Check that the AzureAD module is available
    if (!(Get-Module AzureAD))
    {
        Import-Module AzureAD
    }
    
    #Get username and password of AzureAD service account from encrypted files
    $AADUser = "ADMIN_NAME@DOMAIN.NAME"
    $AADPasswordFile = "\\DOMAIN.NAME\daten\Scripts\Onboarding\Keys\AzureADPwd.txt"
    $AADKeyFile = "\\DOMAIN.NAME\daten\Scripts\Onboarding\Keys\AzureADKey.key"
    $AADKey = Get-Content $AADKeyFile
    $AzureADCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AADUser, (Get-Content $AADPasswordFile | ConvertTo-SecureString -Key $AADKey)
    
    #Connect to AzureAD
    Connect-AzureAD -Credential $AzureADCredential | out-null

    #Connect to Exchange Online
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $AzureADCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking 

    Try
    {
        $users = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled -eq $false}
        foreach ($user in $users)
        {
            Try
            {
                #Write-Host "Checking Azure AD user to see if they need to be removed from their mail distribution groups:"$user.UserPrincipalName -ForegroundColor Cyan
                $UserDistinguishedName = (Get-User $user.UserPrincipalName).DistinguishedName

                #Get the distribution groups membership of each user
                $DistributionGroups = Get-Recipient -Filter "members -eq '$($UserDistinguishedName)'"
                if($DistributionGroups -ne $null)
                {
                    foreach($DistributionGroup in $DistributionGroups)
                    {
                        #Remove the Azure AD user from all distribution groups to which the user belongs
                        Write-Host "Removing user: "$user.UserPrincipalName " from mail distribution group: " $DistributionGroup.Name -ForegroundColor Green
                        Remove-DistributionGroupMember -Identity $DistributionGroup.DistinguishedName -Member $UserDistinguishedName -Confirm:$false -ErrorAction Continue
                    }
                }
                else
                {
                    Write-Host "User either does not have a mailbox or is not in any distribution groups:"$user.UserPrincipalName -ForegroundColor White
                }
            }
            Catch
            {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Could not remove: "$user.UserPrincipalName             -ForegroundColor Red -BackgroundColor Black
                Write-Host "from distribution group: "$DistributionGroup           -ForegroundColor Red -BackgroundColor Black
                Write-Warning -Message $_.Exception.Message
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
            }

            ##################################################################################################################
            Try
            {
                #Write-Host "Checking Azure AD user to see if they need to be removed from any of the shared mailboxes:"$user.UserPrincipalName -ForegroundColor Cyan
                #Collect all shared mailboxes in Office365
                $SharedMailboxs = Get-Mailbox -RecipientTypeDetails SharedMailbox

                #Create a flag variable to determine the console output after looping through all the mailboxes
                $IsASharedMailboxMember = $false
                
                #Loop through the distribution groups
                foreach($SharedMailbox in $SharedMailboxs)
                {
                    #Write-Host "Checking shared mailbox:" $SharedMailbox.name -ForegroundColor Cyan
                    #Collect the members of each shared mailbox
                    $SharedMailboxPermissions = Get-MailboxPermission -Identity $SharedMailbox.DistinguishedName
                    foreach($SharedMailboxPermission in $SharedMailboxPermissions)
                    {
                        #Write-Host "and user member: "$SharedMailboxPermission.User.ToString() -ForegroundColor Cyan
                        #Check if the user is in the distribution group
                        if($SharedMailboxPermission.User.ToString() -eq $user.UserPrincipalName)
                        {
                            #Remove the Azure AD user from all distribution groups to which the user belongs
                            Write-Host "Removing user:"$user.UserPrincipalName" from shared mailbox: "$SharedMailbox.Name -ForegroundColor Green
                            Remove-MailboxPermission -Identity $SharedMailbox.DistinguishedName -User $user.UserPrincipalName -AccessRights FullAccess,SendAs,ExternalAccount,DeleteItem,ReadPermission,ChangePermission,ChangeOwner -InheritanceType All -Confirm:$false -ErrorAction Continue
                            $IsASharedMailboxMember = $true
                        }
                    }
                }
                if ($IsASharedMailboxMember -eq $false)
                {
                    Write-Host "Has no permissions on any shared mailbox:"$user.UserPrincipalName -ForegroundColor White
                }
            }
            Catch
            {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Could not remove: "$user.UserPrincipalName             -ForegroundColor Red -BackgroundColor Black
                Write-Host "from distribution group: "$DistributionGroup           -ForegroundColor Red -BackgroundColor Black
                Write-Warning -Message $_.Exception.Message
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
            }
            Try
            {
                #Write-Host "Checking Azure AD user to see if they need to be removed from any Azure AD groups:" $user.UserPrincipalName -ForegroundColor Cyan
                #Get the Azure AD User's ID
                $userID = (Get-AzureADUser -ObjectId $user.UserPrincipalName).ObjectID
                #Get the Azure groups to which the Azure AD user belongs
                $AzureADGroups = Get-AzureADUserMembership -ObjectId $user.UserPrincipalName
                if ($AzureADGroups -ne $null)
                {
                    foreach ($AzureADGroup in $AzureADGroups)
                    {
                        if(($AzureADGroup.DirSyncEnabled -eq $null) -and ($AzureADGroup.DisplayName -ne "All Users"))
                        {
                            #Remove the Azure AD user from all Azure groups to which the user belongs
                            Write-Host "Removing user:" $user.UserPrincipalName"from Azure AD Group:"$AzureADGroup.DisplayName -ForegroundColor Green
                            Remove-AzureADGroupMember -ObjectId $AzureADGroup.ObjectID -MemberId $userID -ErrorAction Continue
                        }
                        else
                        {
                            Write-Host "Is not in any static Azure AD groups: "$user.UserPrincipalName -ForegroundColor White
                        }
                    }
                }
                else
                {
                    Write-Host "Is not in any Azure AD groups: "$user.UserPrincipalName -ForegroundColor White
                }
            }
            Catch
            {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Could not remove: "$user.UserPrincipalName       -ForegroundColor Red -BackgroundColor Black
                Write-Host "from Azure group: "$AzureADGroup.Name            -ForegroundColor Red -BackgroundColor Black
                Write-Warning -Message $_.Exception.Message
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
            }
            Write-Host "===============================================================================================================================" -ForegroundColor Yellow
        }
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "General failure removing the Azure AD user from their distriubtion groups" -ForegroundColor Red -BackgroundColor Black
        Write-Host "on user:"$user.UserPrincipalName                                           -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
    }
Remove-PSSession $Session
}

#Usage
removeAzureADUserFromAllOffice365Groups