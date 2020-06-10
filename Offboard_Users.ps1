<#
----- PREREQUISITES -----

*You need RSAT to for the Active Directory administration module for Powershell
 Here is a link to the version for Windows 10 x64 1809: https://www.microsoft.com/en-us/download/confirmation.aspx?id=45520

 You also have to install the Sharepoint Online Service Module for Powershell using this command run in Powershell in admin mode
 Install-Module -Name Microsoft.Online.SharePoint.PowerShell -RequiredVersion 16.0.8029.0


----- END PREREQUISITES -----
#>

#####################
### Begin logging ###
#####################
function beginLogging
{
    #Begin logging
    $LogPath = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Logs\Offboarding\"
    $Date = Get-Date -Format "dd-mm-yyyy_HH-mm-ss" 
    $LogFile = $LogPath+'\OffboardUsers_'+$Date+'.log'
    Start-Transcript -Path $LogFile
}

##########################
### Connect to AzureAD ###
##########################
function connectToAzureAD
{
    #Perform the online connections in a separate function to make offline testing easier
    #Connect to Office 365 via AAD Graph
    #You may need to install the module as show below
    #Install-Module -Name AzureAD

    #Check that the AzureAD module is available
    if (!(Get-Module AzureAD))
    {
        Import-Module AzureAD
    }
    
    #Get username and password of AzureAD service account from encrypted files
    $AADUser = "ADMIN@DOMAIN.NAME"
    $AADPasswordFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\AzureADPwd.txt"
    $AADKeyFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\AzureADKey.key"
    $AADKey = Get-Content $AADKeyFile
    $AzureADCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AADUser, (Get-Content $AADPasswordFile | ConvertTo-SecureString -Key $AADKey)
    
    #Connect to AzureAD
    Connect-AzureAD -Credential $AzureADCredential | out-null
    return $AzureADCredential
}

##################################
### Connect to Exchange Online ###
##################################
function connectToExchangeOnline
{
    #Perform the online connections in a separate function to make offline testing easier
    #Connect to Office365 Exchange Online
    
    #Get username and password of AzureAD service account from encrypted files
    $AADUser = "ADMIN@DOMAIN.NAME"
    $AADPasswordFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\AzureADPwd.txt"
    $AADKeyFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\AzureADKey.key"
    $AADKey = Get-Content $AADKeyFile
    $AzureADCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AADUser, (Get-Content $AADPasswordFile | ConvertTo-SecureString -Key $AADKey)

    #Connect to Exchange Online
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $AzureADCredential -Authentication Basic -AllowRedirection
    return $Session
}

#################################################################################################
### Open a file dialog for the user to choose the CSV file with the relevant user information ###
#################################################################################################
function getCSVPath
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $Path = $OpenFileDialog.FileName
    return $Path
}

###################################################################################################
### Generate an array of needed variables for AD/Office365 accounts and puts them into an array ###
###################################################################################################
function generateUserAccountVariables
{
    #$userAccountVariables[0] is the company name as it appears in the OUs in AD
    #$userAccountVariables[1] is the SAMAccountName, i.e. USERNAME
    #$userAccountVariables[2] is the UserPrincipalName i.e. USERNAME@DOMAIN.NAME
    #$userAccountVariables[3] is the user's lastname except for COMPANY_NAME3 users who use their firstname for Azure ADs mailNickName attribute
    #$userAccountVariables[4] is the user who will receive the deprovisioned user's mailbox as a shared mailbox and their OneDrive
    #$userAccountVariables[5] is the company of the user who will receive the deprovisioned user's mailbox and OneDrive

    param
    (
        [string]$company,
        [string]$firstName,
        [string]$lastName,
        [string]$mailRecipient,
        [string]$mailRecipientCompany
    )

    #Clean up the gnarly German names
    $firstname = $firstName.Replace("ä","ae")
    $firstname = $firstName.Replace("ö","oe")
    $firstname = $firstName.Replace("ü","ue")
    $firstname = $firstName.Replace("ß","ss")
    $lastname = $lastname.Replace("ä","ae")
    $lastname = $lastname.Replace("ö","oe")
    $lastname = $lastname.Replace("ü","ue")
    $lastname = $lastname.Replace("ß","ss")

    #Check that the WhoInheritsTheMailAccountsCompany cell from the CSV file contains properly company names
    #i.e. COMPANY_NAME, COMPANY_NAME2, COMPANY_NAME3, COMPANY_NAME3, or COMPANY_NAME4
    if(($mailRecipientCompany -ne $null) -and ($mailRecipientCompany -ne "") -and ($mailRecipientCompany -ne "COMPANY_NAME") -and ($mailRecipientCompany -ne "COMPANY_NAME")`        -and ($mailRecipientCompany -ne "COMPANY_NAME2") -and ($mailRecipientCompany -ne "COMPANY_NAME3") -and ($mailRecipientCompany -ne "COMPANY_NAME3") `        -and ($mailRecipientCompany -ne "COMPANY_NAME4"))
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "Error: The WhoInheritsTheMailAccountsCompany field in the CSV file has an     " -ForegroundColor Red -BackgroundColor Black
        Write-Host "has an invalid company name.  The only accepted values are                    " -ForegroundColor Red -BackgroundColor Black
        Write-Host "COMPANY_NAME, COMPANY_NAME2, COMPANY_NAME3, COMPANY_NAME3, or COMPANY_NAME4   " -ForegroundColor Red -BackgroundColor Black
        Write-Host "Skipping this user:" $UserPrincipalName                                         -ForegroundColor Red -BackgroundColor Black 
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        break
    }

    #Generate User account variables based on the user to be deactivated's company
    if((($company -imatch "COMPANY_NAME") -or ($company -imatch "COMPANY_NAME")) -and ($company -inotmatch "COMPANY_NAME3"))
    {
        $OUName = "COMPANY_NAME_AG"
        $SAMAccountName = $firstName[0]+$lastName
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"
        $mailNickName = $lastName

        #Due to duplicate possible users, get the mail recipient's UPN from local AD based on their name and the company that the deactivated user works for
        if(($mailRecipient -ne $null) -and ($mailRecipient -ne ""))
        {
            $mailRecipientDisplayName = '*'+$mailRecipient+' | '+$mailRecipientCompany+'*'
            $mailRecipient = (Get-ADuser -Filter {displayname -like $mailRecipientDisplayName}).UserPrincipalName
        }
        
        return $OUName, $SAMAccountName, $UserPrincipalName, $mailNickName, $mailRecipient
    }
    elseif($company -imatch "COMPANY_NAME3")
    {
        $OUName = "COMPANY_NAME3"
        $SAMAccountName = $firstName[0]+$lastName
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"
        $mailNickName = $firstName
        
        #Due to duplicate possible users, get the mail recipient's UPN from local AD based on their name and the company that the deactivated user works for
        if(($mailRecipient -ne $null) -and ($mailRecipient -ne ""))
        {
            $mailRecipientDisplayName = '*'+$mailRecipient+' | '+$mailRecipientCompany+'*'
            $mailRecipient = (Get-ADuser -Filter {displayname -like $mailRecipientDisplayName}).UserPrincipalName
        }
        
        return $OUName, $SAMAccountName, $UserPrincipalName, $mailNickName, $mailRecipient
    }
    elseif($company -imatch "COMPANY_NAME2")
    {
        $OUName = "COMPANY_NAME2_GmbH"
        $SAMAccountName = $firstName[0]+$lastName
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"
        $mailNickName = $lastName

        #Due to duplicate possible users, get the mail recipient's UPN from local AD based on their name and the company that the deactivated user works for
        if(($mailRecipient -ne $null) -and ($mailRecipient -ne ""))
        {
            $mailRecipientDisplayName = '*'+$mailRecipient+' | '+$mailRecipientCompany+'*'
            $mailRecipient = (Get-ADuser -Filter {displayname -like $mailRecipientDisplayName}).UserPrincipalName
        }
        
        return $OUName, $SAMAccountName, $UserPrincipalName, $mailNickName, $mailRecipient
    }
    elseif($company -imatch "COMPANY_NAME4")
    {
        $OUName = "COMPANY_NAME_AG"
        $SAMAccountName = $firstName[0]+$lastName
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"
        $mailNickName = $lastName

        #Due to duplicate possible users, get the mail recipient's UPN from local AD based on their name and the company that the deactivated user works for
        if(($mailRecipient -ne $null) -and ($mailRecipient -ne ""))
        {
            $mailRecipientDisplayName = '*'+$mailRecipient+' | '+$mailRecipientCompany+'*'
            $mailRecipient = (Get-ADuser -Filter {displayname -like $mailRecipientDisplayName}).UserPrincipalName
        }
        
        return $OUName, $SAMAccountName, $UserPrincipalName, $mailNickName, $mailRecipient
    }
    elseif($company -imatch "COMPANY_NAME3")
    {
        $OUName = "COMPANY_NAME_AG" 
        $SAMAccountName = $firstName[0]+$lastName
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"
        $mailNickName = $lastName
        
        #Due to duplicate possible users, get the mail recipient's UPN from local AD based on their name and the company that the deactivated user works for
        if(($mailRecipient -ne $null) -and ($mailRecipient -ne ""))
        {
            $mailRecipientDisplayName = '*'+$mailRecipient+' | '+$mailRecipientCompany+'*'
            $mailRecipient = (Get-ADuser -Filter {displayname -like $mailRecipientDisplayName}).UserPrincipalName
        }
        
        return $OUName, $SAMAccountName, $UserPrincipalName, $mailNickName, $mailRecipient
    }
    #If they do not work for the correct company report this and break the loop
    else
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "Error: Company name not accepted. Skipping this user        " -ForegroundColor Red -BackgroundColor Black
        Write-Host "Please check the CSV and ensure that company name is either:" -ForegroundColor Red -BackgroundColor Black
        Write-Host "COMPANY_NAME, COMPANY_NAME3, COMPANY_NAME2 GmbH, COMPANY_NAME4, or COMPANY_NAME4    " -ForegroundColor Red -BackgroundColor Black
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        break
    }
}

#######################################################################
### Deprovision a user in local AD using the info from the CSV file ###
#######################################################################
function deprovisionLocalADUser
{
    param
    (
       [string] $CSVPath
    )

    Write-Host "==============================================" -ForegroundColor Yellow
    Write-Host "Deprovisioning Users in local Active Directory" -ForegroundColor Yellow
    Write-Host "==============================================" -ForegroundColor Yellow
    
	Try 
    {
        #Import the CSV File which is passed as parameter when the script is called - Again, make sure the CSV is ";" delimited and not "," delimited
        $users = Import-CSV -path $CSVPath -Encoding UTF8 -UseCulture

        #Check that the AD module is available
        if (!(Get-Module ActiveDirectory))
        {
            Import-Module ActiveDirectory
        }

        #Get username and password of the local AD user service account from encrypted files
        $ADUser = "DOMAIN.NAME\AD_ADMIN_NAME"
        $ADPasswordFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\ADPwd.txt"
        $ADKeyFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\ADKey.key"
        $ADKey = Get-Content $ADKeyFile
        $ADCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUser, (Get-Content $ADPasswordFile | ConvertTo-SecureString -Key $ADKey)

        #Loop through users in CSV file
        foreach ($user in $users) 
        {
            Write-Host "Deactivating local AD user:"$user.FirstName $user.LastName -ForegroundColor Cyan
            #Run the generateUserAccountVariables function to generate an array of needed variables for AD
            $userAccountVariables = generateUserAccountVariables $user.Company $user.FirstName $user.LastName

            #Generate the path of the local AD user object
            $Path = 'OU=Users,OU='+$userAccountVariables[0]+',DC=COMPANY_NAME,DC=de' #i.e. OU=Users,OU=COMPANY_NAME,DC=COMPANY_NAME,DC=de
            
            #Check that the user does not already exist by comparing SAMAccountNames and the AccountCreated column in the CSV file
            $ADFilter = "userPrincipalName -eq "+''''+$userAccountVariables[2]+''''
            $ADUserToDisable = Get-ADUser -Filter $ADFilter -ErrorAction SilentlyContinue

            if (($user.ADAccountDisabled -ne "") -and ($user.ADAccountDisabled -ine "nein") -and ($user.ADAccountDisabled -ine "no"))
            {
                #If user appears to have already been disabled as indicated by the CSV file, output a warning message
                Write-Host "Warning: a local AD user with the UserPrincipalName: $userAccountVariables[2] appears to have already been disabled as indicated by the CSV file. Skipping..." -ForegroundColor White -BackgroundColor DarkBlue
            }
            elseif ($ADUserToDisable -eq $null)
            {
                #If user does not exist in local AD, output a warning message
                Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: a local AD user with the UserPrincipalName:               " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host  $userAccountVariables[2]" does not exist in a search of local AD.  " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Check for typos or misspelled  names in local AD or the CSV file.  " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }
            else
            {
                Try
                {
                    Write-Host "Moving local AD user to Locked User OU:"$user.FirstName $user.LastName -ForegroundColor Cyan
                    #Disable the AD user object
                    Set-ADUser -Identity $ADUSerToDisable -Enabled $false -Credential $ADCredential

                    #Move the AD user object to the "Locked User" OU
                    $NewPath = 'OU=Locked User,OU=COMPANY_NAME_AG,DC=COMPANY_NAME,DC=de'
                    Move-ADObject -Identity $ADUSerToDisable -TargetPath $NewPath -Credential $ADCredential

                    $user.ADAccountDisabled = "DISABLED"
                    $user.ADAccountDisabledDate = Get-Date -Format "dd-MM-yyyy"
                }
                Catch
                {
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                    Write-Host "Could not disable or move local AD User:"$user.FirstName $user.LastName   -ForegroundColor Red -BackgroundColor Black
                    Write-Warning -Message $_.Exception.Message
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                }
            }
            Write-Host "======================================================================================" -ForegroundColor Yellow
        }# End of for loop for the CSV file
        Write-Host "Finished with looping through local AD users in the Offboarding CSV file" -ForegroundColor Cyan
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "General failure setting up the disable local AD user steps"  -ForegroundColor Red -BackgroundColor Black
        Write-Host "on user:"$user.FirstName $user.LastName                      -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
    }
    #Write to the CSV to indicate that the account was created in local AD
    $users | Export-CSV -Path $CSVPath -UseCulture -Encoding UTF8 -NoTypeInformation
}

#######################################
### Remove user from all ACL Groups ###
#######################################
function removeLocalADUserFromACLGroups
{
    param
    (
       [string] $CSVPath
    )

    Write-Host "=======================================" -ForegroundColor Yellow
    Write-Host "Removing local AD users from ACL Groups" -ForegroundColor Yellow
    Write-Host "=======================================" -ForegroundColor Yellow

    Try 
    {
        #Import the CSV File which is passed as parameter when the script is called - Again, make sure the CSV is ";" delimited and not "," delimited
        $users = Import-CSV -path $CSVPath -Encoding UTF8 -UseCulture

        #Check that the AD module is available
        if (!(Get-Module ActiveDirectory))
        {
            Import-Module ActiveDirectory
        }

        #Get username and password of the local AD user service account from encrypted files
        $ADUser = "DOMAIN.NAME\AD_ADMIN_NAME"
        $ADPasswordFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\ADPwd.txt"
        $ADKeyFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\ADKey.key"
        $ADKey = Get-Content $ADKeyFile
        $ADCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUser, (Get-Content $ADPasswordFile | ConvertTo-SecureString -Key $ADKey)

        #Loop through users in CSV file
        foreach ($user in $users) 
        {
            #Run the generateUserAccountVariables function to generate an array of needed variables for AD
            $userAccountVariables = generateUserAccountVariables $user.Company $user.FirstName $user.LastName

            #Check that the user does not already exist by comparing SAMAccountNames and the AccountCreated column in the CSV file
            $ADFilter = "userPrincipalName -eq "+''''+$userAccountVariables[2]+''''
            $ADUserToRemoveFromACLGroups = Get-ADUser -Filter $ADFilter -ErrorAction SilentlyContinue
            if ($ADUserToRemoveFromACLGroups -eq $null)
            {
                #If user does not exist in local AD, output a warning message
                Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: a local AD user with the UserPrincipalName:               " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host  $userAccountVariables[2]" does not exist in a search of local AD.  " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Check for typos or misspelled  names in local AD or the CSV file.  " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }
            else
            {
                Try
                {
                    #Get the groups to which the local AD user belongs
                    $Groups = Get-ADPrincipalGroupMembership -Identity $ADUserToRemoveFromACLGroups

                    foreach ($Group in $Groups)
                    {
                        if($Group.Name -ne "Domain Users")
                        {
                            #Remove the local AD user from all groups to which the user belongs EXCEPT for Domain Users
                            Write-Host "Removing user: "$user.FirstName $user.LastName " from " $Group.Name -ForegroundColor Cyan
                            Remove-ADGroupMember -Identity $Group -Members $ADUserToRemoveFromACLGroups -Credential $ADCredential -Confirm:$false
                        }
                    }
                }
                Catch
                {
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                    Write-Host "Could not remove local AD User:"$userAccountVariables[2]        -ForegroundColor Red -BackgroundColor Black
                    Write-Host "from their various ACL Groups"                                  -ForegroundColor Red -BackgroundColor Black
                    Write-Warning -Message $_.Exception.Message
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                }
            }
        }
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "General failure removing the local AD User from their ACL groups"  -ForegroundColor Red -BackgroundColor Black
        Write-Host "on user:"$user.FirstName $user.LastName                            -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
    }
}

#######################################################################
### Deprovision a user in Azure AD using the info from the CSV file ###
#######################################################################
function deprovisionAzureADUser
{
    param
    (
       [string] $CSVPath
    )

    Write-Host "==============================================" -ForegroundColor Yellow
    Write-Host "Deprovisioning Users in Azure Active Directory" -ForegroundColor Yellow
    Write-Host "==============================================" -ForegroundColor Yellow
    
	Try 
    {
        #Connect to Azure AD
        $AzureADCredential = connectToAzureAD

        #Connect to Exchange Online
        $Session = connectToExchangeOnline
        Import-PSSession $Session -DisableNameChecking 

        #Import the CSV File which is passed as parameter when the script is called - Again, make sure the CSV is ";" delimited and not "," delimited
        $users = Import-CSV -path $CSVPath -Encoding UTF8 -UseCulture

        #Loop through users in CSV file
        foreach ($user in $users) 
        {
            #Run the generateUserAccountVariables function to generate an array of needed variables for Azure AD
            $userAccountVariables = generateUserAccountVariables $user.Company $user.FirstName $user.LastName $user.WhoInheritsTheMailAccount $user.WhoInheritsTheMailAccountsCompany
            Write-Host "Deactivating Azure AD user:"$userAccountVariables[2] -ForegroundColor Cyan

            #Check that the user does exist by comparing SAMAccountNames and the AccountCreated column in the CSV file
            $AzureADFilter = "userPrincipalName eq "+''''+$userAccountVariables[2]+''''
            $AzureADUserToDisable = Get-AzureADUser -Filter $AzureADFilter
            $AzureADUserMailbox = Get-Mailbox -Identity $userAccountVariables[2]

            #Check that the Mail Recipient (WhoInheritsTheMailAccount from the csv) exists and has a mailbox          
            if(($userAccountVariables[4] -ne $null) -and ($userAccountVariables[4] -ne $null))
            {
                $AzureADUserMailRecipientFilter = "userPrincipalName eq "+''''+$userAccountVariables[4]+''''
                $AzureADUserMailRecipient = Get-AzureADUser -Filter $AzureADADUserMailRecipientFilter
                $AzureADUserMailboxMailRecipient = Get-Mailbox -Identity $userAccountVariables[4]
            }

            if (($user.AzureADAccountDisabled -ne "") -and ($user.AzureADAccountDisabled -ine "nein") -and ($user.AzureADAccountDisabled -ine "no"))
            {
                #If user appears to have already been disabled as indicated by the CSV file, output a warning message
                Write-Host "Warning: an Azure AD user with the UserPrincipalName:" $userAccountVariables[2] "appears to have already been disabled as indicated by the CSV file. Skipping..." -ForegroundColor White -BackgroundColor DarkBlue
            }
            elseif (($AzureADUserToDisable -eq $null) -and ($AzureADUserMailbox -ne $null))
            {
                #If user does not exist in Azure AD, output a warning message
                Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: an Azure AD user with the UserPrincipalName:              " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host  $userAccountVariables[2]" does not exist in a search of Azure AD.  " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Check for typos or misspelled names in Azure AD or the CSV file.   " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }
            elseif (($userAccountVariables[4] -eq "@DOMAIN.NAME") -or ($AzureADUserMailRecipient -eq $null) -or ($AzureADUserMailboxMailRecipient -eq $null))
            {
                #If the mail recipient user does not exist in Azure AD, output a warning message
                Write-Host "====================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: an Azure AD Mail Recipient user with the UserPrincipalName:" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host  $userAccountVariables[4]" does not exist in a search of Azure AD.   " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Check for typos or misspelled names in Azure AD or the CSV file.    " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "====================================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }
            else
            {
                Try
                {                    
                    #Disable the AD user object
                    Set-AzureADUser -ObjectID $AzureADUserToDisable.ObjectID -AccountEnabled $false -ErrorAction SilentlyContinue

                    #Convert the mailbox of the Azure AD user to a shared mailbox
                    Set-Mailbox -Identity $userAccountVariables[2] -Type Shared -WarningAction silentlyContinue

                    #Take a little break
                    Start-Sleep -s 5

                    #Assign full access rights user from the CSV's WhoInheritsTheMailAccount field to the Shared Mailbox, unless it is empty then produce a warning
                    if (($userAccountVariables[4] -ne "@DOMAIN.NAME") -and ($AzureADUserMailRecipient -ne $null) -and ($AzureADUserMailboxMailRecipient -ne $null))
                    {
                        Write-Host "Assigning Shared Mailbox access rights to"$userAccountVariables[4]"for the mailbox of former user"$user.FirstName $user.LastName -ForegroundColor Cyan
                        Add-MailboxPermission -Identity $userAccountVariables[2] -User $userAccountVariables[4] -AccessRights FullAccess -InheritanceType All -WarningAction silentlyContinue | Out-Null
                    }
                    else
                    {
                        Write-Host "==========================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                        Write-Host "Warning: the WhoInheritsTheMailAccount field in the CSV file is empty for:" -ForegroundColor White -BackgroundColor DarkBlue
                        Write-Host  $userAccountVariables[2]" thus no one will be assigned rights to their    " -ForegroundColor White -BackgroundColor DarkBlue
                        Write-Host "Exchange mailbox or OneDrive account.  You might want to check this...    " -ForegroundColor White -BackgroundColor DarkBlue
                        Write-Host "==========================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                    }

                    #Assign full access rights to the user from the CSV's WhoInheritsTheMailAccount field to disabled user's OneDrive
                    if (($userAccountVariables[4] -ne "@DOMAIN.NAME") -and ($AzureADUserMailRecipient -ne $null) -and ($AzureADUserMailboxMailRecipient -ne $null))
                    {
                        Write-Host "Assigning admin access rights to"$userAccountVariables[4]"for the OneDrive account of former user"$user.FirstName $user.LastName -ForegroundColor Cyan
                        Connect-SPOService -Url https://COMPANY_NAMEgruppe-admin.sharepoint.com -credential $AzureADCredential | Out-Null
                        $OneDriveSite = "https://COMPANY_NAMEgruppe-my.sharepoint.com/personal/"+$userAccountVariables[1]+"_COMPANY_NAME_de/_layouts/15/onedrive.aspx"
                        Set-SPOUser -Site $OneDriveSite -LoginName $userAccountVariables[4] -IsSiteCollectionAdmin $true | Out-Null
                        Disconnect-SPOService | Out-Null
                    }

                    #Check if the user has any Office365 licenses and remove them
                    $licensesToRemove = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                    $licenses = (Get-AzureADUser -ObjectID $userAccountVariables[2]).AssignedLicenses
                    if($licenses -ne $null)
                    {
                        foreach($license in $licenses)
                        {
                            $licensesToRemove.RemoveLicenses = $license.SkuId
                            Write-Host "Removed Office365 license with SkuID of:"$license.SkuId"from"$userAccountVariables[2] -ForegroundColor Cyan

                        }
                        Set-AzureADUserLicense -ObjectId $userAccountVariables[2] -AssignedLicenses $licensesToRemove -ErrorAction Continue
                    }
                    else
                    {
                        #Output a warning message if the user has no assigned Office365 licenses
                        Write-Host "==========================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                        Write-Host "Warning: an Azure AD user with the UserPrincipalName:                     " -ForegroundColor White -BackgroundColor DarkBlue
                        Write-Host  $userAccountVariables[2]" had no assigned Office365 licenses.  Skipping..." -ForegroundColor White -BackgroundColor DarkBlue
                        Write-Host "==========================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                    }

                    #Mark the account as disabled in the CSV so that it is skipped next run
                    $user.AzureADAccountDisabled = "DISABLED"
                    $user.AzureADAccountDisabledDate = Get-Date -Format "dd-MM-yyyy"
                }
                Catch
                {
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                    Write-Host "Could not disable or move Azure AD User:"$userAccountVariables[2]         -ForegroundColor Red -BackgroundColor Black
                    Write-Warning -Message $_.Exception.Message
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black

                    #Don't forget to disconnect your exchange online session
                    Remove-PSSession $Session
                }
            }
            Write-Host "======================================================================================" -ForegroundColor Yellow
        }# End of for loop for the CSV file
        Write-Host "Finished with looping through Azure AD users in the Offboarding CSV file" -ForegroundColor Cyan
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "General failure setting up the disable Azure AD user steps"  -ForegroundColor Red -BackgroundColor Black
        Write-Host "on user:"$user.FirstName $user.LastName                      -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black

        #Don't forget to disconnect your exchange online session
        Remove-PSSession $Session
    }

    #Write to the CSV to indicate that the account was created in Azure AD
    $users | Export-CSV -Path $CSVPath -UseCulture -Encoding UTF8 -NoTypeInformation

    #Don't forget to disconnect your exchange online session
    Remove-PSSession $Session
}

#############################################################
### Remove Azure AD users from all Office365/Azure Groups ###
#############################################################
function removeAzureADUserFromAllOffice365Groups
{
    param
    (
       [string] $CSVPath
    )

    Write-Host "===========================================================================" -ForegroundColor Yellow
    Write-Host "Removing deactivated Azure AD Users from all Office365 Groups and Mailboxes" -ForegroundColor Yellow
    Write-Host "===========================================================================" -ForegroundColor Yellow
    
    #Connect to Azure AD
    connectToAzureAD

    #Connect to Exchange Online
    $Session = connectToExchangeOnline
    Import-PSSession $Session -DisableNameChecking 

    Try
    {
        #Import the CSV File which is passed as parameter when the script is called - Again, make sure the CSV is ";" delimited and not "," delimited
        $users = Import-CSV -path $CSVPath -Encoding UTF8 -UseCulture
        foreach ($user in $users)
        {
            Try
            {
                Write-Host "Checking Azure AD user to see if they need to be removed from their mail distribution groups:"$user.UserPrincipalName -ForegroundColor Cyan
                #Run the generateUserAccountVariables function to generate an array of needed variables for Azure AD
                $userAccountVariables = generateUserAccountVariables $user.Company $user.FirstName $user.LastName $user.WhoInheritsTheMailAccount
                
                $UserDistinguishedName = (Get-User $userAccountVariables[2]).DistinguishedName

                #Get the distribution groups membership of each user
                $DistributionGroups = Get-Recipient -Filter "members -eq '$($UserDistinguishedName)'"
                if($DistributionGroups -ne $null)
                {
                    foreach($DistributionGroup in $DistributionGroups)
                    {
                        if($DistributionGroup.RecipientTypeDetails -eq "MailUniversalDistributionGroup")
                        {
                            #Remove the Azure AD user from all distribution groups to which the user belongs
                            Write-Host "Removing user: "$userAccountVariables[2] " from mail distribution group: " $DistributionGroup.Name -ForegroundColor Cyan
                            Remove-DistributionGroupMember -Identity $DistributionGroup.DistinguishedName -Member $UserDistinguishedName -Confirm:$false -ErrorAction Continue
                        }
                        elseif($DistributionGroup.RecipientTypeDetails -eq "GroupMailbox")
                        {
                            #Remove the Azure AD user from all Office365 groups to which the user belongs
                            Write-Host "Removing user: "$userAccountVariables[2] " from Office365 group: " $DistributionGroup.Name -ForegroundColor Cyan
                            Remove-UnifiedGroupLinks -Identity $DistributionGroup.DistinguishedName -LinkType Members -Links $UserDistinguishedName -Confirm:$false -ErrorAction Continue
                        }
                        else
                        {
                            #If the group type is not MailUniversalDistributionGroup or GroupMailbox then output a warning message
                            Write-Host "=============================================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                            Write-Host "Warning: an Azure AD user with the UserPrincipalName:                                        " -ForegroundColor White -BackgroundColor DarkBlue
                            Write-Host  $userAccountVariables[2]" is a member of the group:                                          " -ForegroundColor White -BackgroundColor DarkBlue
                            Write-Host  $DistributionGroup.Alias "which is a"$DistributionGroup.RecipientTypeDetailsCheck"type group." -ForegroundColor White -BackgroundColor DarkBlue
                            Write-Host "The script does not know how to remvove a user from this type of group...                    " -ForegroundColor White -BackgroundColor DarkBlue
                            Write-Host "=============================================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                        }
                    }
                }
                else
                {
                    Write-Host "User either does not have a mailbox or is not in any distribution groups:"$userAccountVariables[2] -ForegroundColor White
                }
            }
            Catch
            {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Could not remove: "$userAccountVariables[2]            -ForegroundColor Red -BackgroundColor Black
                Write-Host "from distribution group: "$DistributionGroup           -ForegroundColor Red -BackgroundColor Black
                Write-Warning -Message $_.Exception.Message
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
            }
            #Now check every shared mailbox to see if the user is in there and remove them
            Try
            {
                #Write-Host "Checking Azure AD user to see if they need to be removed from any of the shared mailboxes:"$userAccountVariables[2] -ForegroundColor Cyan
                #Collect all shared mailboxes in Office365
                $SharedMailboxs = Get-Mailbox -RecipientTypeDetails SharedMailbox

                #Create a flag variable to determine the console output after looping through all the mailboxes
                $IsASharedMailboxMember = $false
                
                #Loop through the distribution groups
                foreach($SharedMailbox in $SharedMailboxs)
                {
                    #Collect the members of each shared mailbox
                    $SharedMailboxPermissions = Get-MailboxPermission -Identity $SharedMailbox.DistinguishedName
                    foreach($SharedMailboxPermission in $SharedMailboxPermissions)
                    {
                        #Check if the user is in the distribution group
                        if($SharedMailboxPermission.User.ToString() -eq $userAccountVariables[2])
                        {
                            #Remove the Azure AD user from all distribution groups to which the user belongs
                            Write-Host "Removing user:"$userAccountVariables[2]" from shared mailbox: "$SharedMailbox.Name -ForegroundColor Cyan
                            Remove-MailboxPermission -Identity $SharedMailbox.DistinguishedName -User $userAccountVariables[2] -AccessRights FullAccess, SendAs,ExternalAccount,DeleteItem,ReadPermission,ChangePermission,ChangeOwner -InheritanceType All -Confirm:$false -ErrorAction Continue
                            $IsASharedMailboxMember = $true
                        }
                    }
                }
                if ($IsASharedMailboxMember -eq $false)
                {
                    Write-Host "Has no permissions on any shared mailbox:"$userAccountVariables[2] -ForegroundColor White
                }
            }
            Catch
            {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Could not remove: "$userAccountVariables[2]            -ForegroundColor Red -BackgroundColor Black
                Write-Host "from distribution group: "$DistributionGroup           -ForegroundColor Red -BackgroundColor Black
                Write-Warning -Message $_.Exception.Message
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
            }

            #Now check the user's Azure AD group membership and remove them from all groups
            Try
            {
                Write-Host "Checking Azure AD user to see if they need to be removed from any Azure AD groups:" $user.UserPrincipalName -ForegroundColor Cyan
                #Get the Azure AD User's ID
                $userID = (Get-AzureADUser -ObjectId $userAccountVariables[2]).ObjectID
                #Get the Azure groups to which the Azure AD user belongs
                $AzureADGroups = Get-AzureADUserMembership -ObjectId $userAccountVariables[2]
                if ($AzureADGroups -ne $null)
                {
                    foreach ($AzureADGroup in $AzureADGroups)
                    {
                        if(($AzureADGroup.DirSyncEnabled -eq $null) -and ($AzureADGroup.DisplayName -ne "All Users"))
                        {
                            #Remove the Azure AD user from all Azure groups to which the user belongs
                            Write-Host "Removing user:"$userAccountVariables[2]"from Azure AD Group:"$AzureADGroup.DisplayName -ForegroundColor Cyan
                            Remove-AzureADGroupMember -ObjectId $AzureADGroup.ObjectID -MemberId $userID -ErrorAction Continue
                        }
                        else
                        {
                            Write-Host "Is not in any static Azure AD groups: "$userAccountVariables[2] -ForegroundColor White
                        }
                    }
                }
                else
                {
                    Write-Host "Is not in any Azure AD groups: "$userAccountVariables[2] -ForegroundColor White
                }
            }
            Catch
            {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Could not remove: "$userAccountVariables[2]      -ForegroundColor Red -BackgroundColor Black
                Write-Host "from Azure group: "$AzureADGroup.DisplayName     -ForegroundColor Red -BackgroundColor Black
                Write-Warning -Message $_.Exception.Message
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                #Don't forget to disconnect your exchange online session
                Remove-PSSession $Session
            }
            Write-Host "===============================================================================================================================" -ForegroundColor Yellow
        }
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "General failure removing the Azure AD user from their distriubtion groups" -ForegroundColor Red -BackgroundColor Black
        Write-Host "on user:"$userAccountVariables[2]                                          -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        
        #Don't forget to disconnect your exchange online session
        Remove-PSSession $Session
    }
    #Don't forget to disconnect your exchange online session
    Remove-PSSession $Session
}

#########################################################################
### Delete all log files from this script that are older than 90 days ###
#########################################################################

function deleteOldLogs
{
$LogPath = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Logs\Offboarding\"
$Daysback = "-90"
 
$CurrentDate = Get-Date
$DatetoDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $LogPath | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
}

#################################
### What to do with OneDrive? ###
#################################
#https://docs.microsoft.com/en-us/onedrive/retention-and-deletion?redirectSourcePath=%252farticle%252fef883c48-332c-42f5-8aea-f0e2366c15f9
#Their manager should automatically get the content and I have set up the itsupport@COMPANY_NAME.ag to be the backup admin for when they don't have a manager.

#########################################################################
### Disconnect from open services and close running transcription log ###
#########################################################################
function endLogging
{
    # Discconect from AAD
    Disconnect-AzureAD

    # Write the log file
    Stop-Transcript
}

##############################
### Run the functions here ###
##############################

#Start the transcription log which are saved here: \\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Logs
beginLogging

#Create a path variable for the CSV file from a pop-up dialog using the getCSVPath function
$CSVPath = getCSVPath

#Disable the local AD Account and move to the Locked Users OU
#deprovisionLocalADUser($CSVPath)

#Get all the groups which the user is a member of and remove the user from those groups
#removeLocalADUserFromACLGroups($CSVPath)

#Disable the user's Azure AD/Office 365 account, convert their mailbox to a shared mailbox, assign the mailbox to the WhoInheritsTheMailAccount entry in the CSV, and remove their Office 365 license(s)
deprovisionAzureADUser($CSVPath)

#Get the mail distribution groups, the shared mailboxes, and the Azure AD groups to which the user belongs and then remove them
#removeAzureADUserFromAllOffice365Groups($CSVPath)

#Clean up log files older than 90 days
#deleteOldLogs

#Stop the transcript and save
endLogging