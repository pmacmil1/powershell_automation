<#
----- PREREQUISITES -----

*You need RSAT to for the Active Directory administration module for Powershell
 Here is a link to the version for Windows 10 x64 1809: https://www.microsoft.com/en-us/download/confirmation.aspx?id=45520

----- END PREREQUISITES -----
#>

#####################
### Begin logging ###
#####################
function beginLogging
{
    #Begin logging
    $LogPath = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Logs\Onboarding\"
    $Date = Get-Date -Format "dd-mm-yyyy_HH-mm-ss" 
    $LogFile = $LogPath+'\OnboardUsers_'+$Date+'.log'
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
    $AADUser = "admin.name@DOMAIN.NAME"
    $AADPasswordFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\AzureADPwd.txt"
    $AADKeyFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\AzureADKey.key"
    $AADKey = Get-Content $AADKeyFile
    $AzureADCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AADUser, (Get-Content $AADPasswordFile | ConvertTo-SecureString -Key $AADKey)
    
    #Connect to AzureAD
    Connect-AzureAD -Credential $AzureADCredential | out-null
}

##################################
### Connect to Exchange Online ###
##################################
function connectToExchangeOnline
{
    #Perform the online connections in a separate function to make offline testing easier
    #Connect to Office365 Exchange Online
    
    #Get username and password of AzureAD service account from encrypted files
    $AADUser = "admin.name@DOMAIN.NAME"
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

##########################################################################
### Check that the generated or input email address are not duplicates ###
##########################################################################
function checkEmailAddresses
{
    param
    (
        [string]$emailAddress
    )

    #Check that the mail address is not duplicated and is valid
    $CurrentMailAddresses = Get-ADUser -Filter "*" -Properties EmailAddress | Select EmailAddress | Where Emailaddress -ne $null
    if(($emailAddress -ne $null) -and ($emailAddress -ne ""))
    {
        foreach($CurrentMailAddresse in $CurrentMailAddresses)
        {
            if($emailAddress -eq $CurrentMailAddresse)
            {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Error: Email address is duplicated. Skipping this user:" $emailAddress                          -ForegroundColor Red -BackgroundColor Black
                Write-Host "Please check the CSV and ensure that the email address field                                  " -ForegroundColor Red -BackgroundColor Black 
                Write-Host "has either a valid, non-duplicated mail address or AUTOMATICALLY                              " -ForegroundColor Red -BackgroundColor Black 
                Write-Host "so that the mail address would be generated from the users first and lastname                 " -ForegroundColor Red -BackgroundColor Black 
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                return $false
            }
        }
    }
    else
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "Error: Email address field is empty in the CSV. Skipping this user:" $emailAddress                    -ForegroundColor Red -BackgroundColor Black
        Write-Host "Please check the CSV and ensure that the email address field                                        " -ForegroundColor Red -BackgroundColor Black 
        Write-Host "has either a valid, non-duplicated mail address or the word AUTOMATICALLY                           " -ForegroundColor Red -BackgroundColor Black 
        Write-Host "so that the mail address would be generated from the users first and lastname                       " -ForegroundColor Red -BackgroundColor Black 
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        return $false
    }
}

########################################################################
### Check that the generated User Principal Names are not duplicates ###
########################################################################
function checkUPNs
{
    param
    (
        [string]$UserPrincipalName
    )

    #Check that the User Principal Name is not duplicated and is valid
    $CurrentUPNs = Get-ADUser -Filter "*" -Properties UserPrincipalName | Select UserPrincipalName | Where UserPrincipalName -ne $null
    if(($UserPrincipalName -ne $null) -and ($UserPrincipalName -ne ""))
    {
        foreach($CurrentUPN in $CurrentUPN)
        {
            
            if($UserPrincipalName -eq $CurrentUPN)
            {
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Error: User Principal Name is duplicated. Skipping this user:" $UserPrincipalName                          -ForegroundColor Red -BackgroundColor Black
                Write-Host "Please check the CSV and ensure that the first and last name fields are not duplicated with another user " -ForegroundColor Red -BackgroundColor Black 
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black 
                return $false
            }
        }
    }
    else
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "Error: Either the first or last name fields are empty in the CSV.             " -ForegroundColor Red -BackgroundColor Black
        Write-Host "Skipping this user:" $UserPrincipalName                                         -ForegroundColor Red -BackgroundColor Black 
        Write-Host "Please check that the first and last name fields in the CSV have valid content" -ForegroundColor Red -BackgroundColor Black 
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black 
        return $false
    }
}

########################################################################################################################
### Generate an array of needed variables for AD/Office365 accounts and puts them into an userAccountVariables array ###
########################################################################################################################
function generateUserAccountVariables
{
    #$userAccountVariables[0] is the password
    #$userAccountVariables[1] is the company name as it should appear in the display name
    #$userAccountVariables[2] is the company name as it appears in the OUs in AD
    #$userAccountVariables[3] is the complete display name
    #$userAccountVariables[4] is the SAMAccountName
    #$userAccountVariables[5] is the UserPrincipalName
    #$userAccountVariables[6] is the user's lastname except for COMPANY_NAME users who use their firstname for Azure ADs mailNickName attribute
    #$userAccountVariables[7] is the user's customized mail address
    #$userAccountVariables[8] is the company's homepage
    #$userAccountVariables[9] is the address of the user's office
    #$userAccountVariables[10] is the postal code of the user's office
    #$userAccountVariables[11] is the central phone number for the user's office
    #$userAccountVariables[12] is the central fax number for the user's office

    param
    (
        [string]$company, 
        [string]$firstName, 
        [string]$lastName,
        [string]$emailAddressInput
    )

    $cleanFirstname = $firstName.Replace("ä","ae")
    $cleanFirstname = $firstName.Replace("ö","oe")
    $cleanFirstname = $firstName.Replace("ü","ue")
    $cleanFirstname = $firstName.Replace("ß","ss")
    $cleanLastname = $lastname.Replace("ä","ae")
    $cleanLastname = $lastname.Replace("ö","oe")
    $cleanLastname = $lastname.Replace("ü","ue")
    $cleanLastname = $lastname.Replace("ß","ss")

    #Now begin generated the various variables for the different companies
    if((($company -imatch "COMPANY_NAME") -or ($company -imatch "COMPANY_NAME1")) -and ($company -inotmatch "COMPANY_NAME2"))
    {
        $password = "XXXXXXXX" | ConvertTo-SecureString -AsPlainText -force
        $cleanCompanyName = "COMPANY_NAME"
        $OUName = "COMPANY_NAME"
        $displayName = $firstname + ' ' + $lastname + ' | ' + "COMPANY_NAME"
        $SAMAccountName = $cleanFirstname[0]+$cleanLastname
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"
        
        #Check that the generated User Principal Name using the checkUPNs function to ensure that it is valid and nonduplicated
        $validUserPrincipalName = checkUPNs $UserPrincipalName
        if($validUserPrincipalName -eq $false)
        {
            break  
        }

        $mailNickName = $cleanLastname

        #Check from the CSV if the e-mail address should be generated "automatically" or if it should be set to the content of the email address field in the CSV
        #And check if the e-mail address is valid and non-duplicated using the checkEmailAddresses function
        if ($emailAddressInput -inotmatch "AUTOMATICALLY")
        {
            $EmailAddress = $emailAddressInput
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }
        else
        {
            $EmailAddress = $cleanLastname+"@company.name"
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }
        $HomePage = "HOME_PAGE"
        $StreetAddress = "STREET_ADDRESS"
        $PostalCode = "XXXXX"
        $OfficePhone = "+49 XXXXXXXXXXX"
        $Fax = "+49 XXXXXXXXXX"

        return $password, $cleanCompanyName, $OUName, $displayName, $SAMAccountName, $UserPrincipalName, `
        $mailNickName, $EmailAddress, $HomePage, $StreetAddress, $PostalCode, $OfficePhone, $Fax
    }
    elseif($company -imatch "COMPANY_NAME1")
    {
        $password = "XXXXXXXX" | ConvertTo-SecureString -AsPlainText -force
        $cleanCompanyName = "COMPANY_NAME1"
        $OUName = "COMPANY_NAME1"
        $displayName = $firstName + ' ' + $lastName + ' | ' + "COMPANY_NAME1"
        $SAMAccountName = $cleanFirstname[0]+$cleanLastname
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"

        #Check that the generated User Principal Name using the checkUPNs function to ensure that it is valid and nonduplicated
        $validUserPrincipalName = checkUPNs $UserPrincipalName
        if($validUserPrincipalName -eq $false)
        {
            break  
        }

        $mailNickName = $cleanFirstname

        #Check from the CSV if the e-mail address should be generated "automatically" or if it should be set to the content of the email address field in the CSV
        #And check if the e-mail address is valid and non-duplicated using the checkEmailAddresses function
        if ($emailAddressInput -inotmatch "AUTOMATICALLY")
        {
            $EmailAddress = $emailAddressInput
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }
        else
        {
            $EmailAddress = $cleanFirstname+"@COMPANY_NAME1.com" 
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }
        $HomePage = "www.COMPANY_NAME1.com"
        $StreetAddress = "STREET_ADDRESS"
        $PostalCode = "XXXXX"
        $OfficePhone = "+49 XXXXXXXXXX"
        $Fax = "+49 XXXXXXXXXX"

        return $password, $cleanCompanyName, $OUName, $displayName, $SAMAccountName, $UserPrincipalName, `
        $mailNickName, $EmailAddress, $HomePage, $StreetAddress, $PostalCode, $OfficePhone, $Fax
    }
    elseif($company -imatch "COMPANY_NAME3")
    {
        $password = "XXXXXXX" | ConvertTo-SecureString -AsPlainText -force
        $cleanCompanyName = "COMPANY_NAME3"
        $OUName = "COMPANY_NAME3"
        $displayName = $firstName + ' ' + $lastName + ' | ' + "COMPANY_NAME3"
        $SAMAccountName = $cleanFirstname[0]+$cleanLastname
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"

        #Check that the generated User Principal Name using the checkUPNs function to ensure that it is valid and nonduplicated
        $validUserPrincipalName = checkUPNs $UserPrincipalName
        if($validUserPrincipalName -eq $false)
        {
            break  
        }

        $mailNickName = $cleanLastname

        #Check from the CSV if the e-mail address should be generated "automatically" or if it should be set to the content of the email address field in the CSV
        #And check if the e-mail address is valid and non-duplicated using the checkEmailAddresses function
        if ($emailAddressInput -inotmatch "AUTOMATICALLY")
        {
            $EmailAddress = $emailAddressInput
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }
        else
        {
            $EmailAddress = $cleanLastname+"@COMPANY_NAME3.com"
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }

        $HomePage = "www.COMPANY_NAME4.com"
        $StreetAddress = "STREET_ADDRESS"
        $PostalCode = "XXXXX"
        $OfficePhone = "+49 XXXXXXXXX"
        $Fax = "+49 XXXXXXXXXX"

        return $password, $cleanCompanyName, $OUName, $displayName, $SAMAccountName, $UserPrincipalName, `
        $mailNickName, $EmailAddress, $HomePage, $StreetAddress, $PostalCode, $OfficePhone, $Fax
    }
    elseif($company -imatch "COMPANY_NAME4")
    {
        $password = "XXXXXXXX" | ConvertTo-SecureString -AsPlainText -force
        $cleanCompanyName = "COMPANY_NAME4"
        $OUName = "COMPANY_NAME4"
        $displayName = $firstName + ' ' + $lastName + ' | ' + 'COMPANY_NAME4'
        $SAMAccountName = $cleanFirstname[0]+$cleanLastname
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"

        #Check that the generated User Principal Name using the checkUPNs function to ensure that it is valid and nonduplicated
        $validUserPrincipalName = checkUPNs $UserPrincipalName
        if($validUserPrincipalName -eq $false)
        {
            break  
        }

        $mailNickName = $cleanLastname

        #Check from the CSV if the e-mail address should be generated "automatically" or if it should be set to the content of the email address field in the CSV
        #And check if the e-mail address is valid and non-duplicated using the checkEmailAddresses function
        if ($emailAddressInput -inotmatch "AUTOMATICALLY")
        {
            $EmailAddress = $emailAddressInput
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }
        else
        {
            $EmailAddress = $cleanLastname+"@company_name4.com"
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }

        $HomePage = "www.COMPANY_NAME5.com"
        $StreetAddress = "STREET_ADDRESS"
        $PostalCode = "XXXXX"
        $OfficePhone = "+49 XXXXXXXXX"
        $Fax = "+49 XXXXXXXXXX"

        return $password, $cleanCompanyName, $OUName, $displayName, $SAMAccountName, $UserPrincipalName, `
        $mailNickName, $EmailAddress, $HomePage, $StreetAddress, $PostalCode, $OfficePhone, $Fax
    }
    elseif($company -imatch "COMPANY_NAME5")
    {
        $password = "XXXXXXXXX" | ConvertTo-SecureString -AsPlainText -force
        $cleanCompanyName = "COMPANY_NAME5"
        $OUName = "COMPANY_NAME5" 
        $displayName = $firstName + ' ' + $lastName + ' | ' + 'COMPANY_NAME5'
        $SAMAccountName = $cleanFirstname[0]+$cleanLastname
        $UserPrincipalName = $SAMAccountName+"@DOMAIN.NAME"

        #Check that the generated User Principal Name using the checkUPNs function to ensure that it is valid and nonduplicated
        $validUserPrincipalName = checkUPNs $UserPrincipalName
        if($validUserPrincipalName -eq $false)
        {
            break  
        }

        $mailNickName = $cleanLastname

        #Check from the CSV if the e-mail address should be generated "automatically" or if it should be set to the content of the email address field in the CSV
        #And check if the e-mail address is valid and non-duplicated using the checkEmailAddresses function
        if ($emailAddressInput -inotmatch "AUTOMATICALLY")
        {
            $EmailAddress = $emailAddressInput
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }
        else
        {
            $EmailAddress = $cleanFirstname[0]+$cleanLastname[0]+"@COMPANY.NAME5"
            $validEmailAddress = checkEmailAddresses $EmailAddress
            if($validEmailAddress -eq $false)
            {
                break  
            }
        }

        $HomePage = "www.COMPANY_NAME5.com"
        $StreetAddress = "STREET_ADDRESS"
        $PostalCode = "XXXXX"
        $OfficePhone = "+49 XXXXXXXXXX"
        $Fax = "+49 XXXXXXXXXX"

        return $password, $cleanCompanyName, $OUName, $displayName, $SAMAccountName, $UserPrincipalName, `
        $mailNickName, $EmailAddress, $HomePage, $StreetAddress, $PostalCode, $OfficePhone, $Fax
    }
    #If they do not work for the correct company report this and break the loop
    else
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "Error: Company name not accepted. Skipping this user        " -ForegroundColor Red -BackgroundColor Black
        Write-Host "Please check the CSV and ensure that company name is either:" -ForegroundColor Red -BackgroundColor Black 
        Write-Host "COMPANY_NAME, COMPANY_NAME1, COMPANY_NAME3, or COMPANY_NAME4" -ForegroundColor Red -BackgroundColor Black 
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black 
        break
    }
}

##################################################################
### Create a user in local AD using the info from the CSV file ###
##################################################################
function createLocalADUser
{
    param
    (
       [string] $CSVPath
    )

    Write-Host "============================================" -ForegroundColor Yellow
    Write-Host "Provisioning Users in local Active Directory" -ForegroundColor Yellow
    Write-Host "============================================" -ForegroundColor Yellow
    
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
        #This account is a member of adm-ad-User-admin, a group which has write to certain OUs, such as the User OUs and the ACL OU
        $ADUser = "DOMAIN.NAME\AD_ADMIN"
        $ADPasswordFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\ADPwd.txt"
        $ADKeyFile = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Keys\ADKey.key"
        $ADKey = Get-Content $ADKeyFile
        $ADCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUser, (Get-Content $ADPasswordFile | ConvertTo-SecureString -Key $ADKey)

        #Loop through users in CSV file
        foreach ($user in $users) 
        {
            Write-Host "Creating local AD user:"$user.FirstName $user.LastName -ForegroundColor Cyan
            #Run the generateUserAccountVariables function to generate an array of needed variables for AD
            $userAccountVariables = generateUserAccountVariables $user.Company $user.FirstName $user.LastName $user.EmailAddress

            #Generate some AD Object specific variables here such as the path, naming convention, UPN, etc...
            $ADObjectName = $user.LastName+", "+$user.FirstName
            $Path = 'OU=User,OU='+$userAccountVariables[2]+',DC=company,DC=com'
            $HomeDirectory = ('\\DOMAIN.NAME\homes$\'+$userAccountVariables[4]).ToLower()
            
            #Check that the user does not already exist by comparing SAMAccountNames and the AccountCreated column in the CSV file
            $ADFilter = "userPrincipalName -eq "+''''+$userAccountVariables[5]+''''
            $DoesADUserExist = Get-ADUser -Filter $ADFilter -ErrorAction SilentlyContinue
            if (($user.ADAccountCreated -ne "") -and ($user.ADAccountCreated -ine "nein") -and ($user.ADAccountCreated -ine "no"))
            {
                #If user appears to exist as indicated by the CSV file, output a warning message
                Write-Host "Warning: a local AD user with the UserPrincipalName: $userAccountVariables[5] appears to exist as indicated by the CSV file. Skipping..." -ForegroundColor White -BackgroundColor DarkBlue
            }
            elseif ($DoesADUserExist -ne $null)
            {
                #If user does exist in local AD, output a warning message
                Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: a local AD user with the UserPrincipalName:               " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host  $userAccountVariables[5]" already exists in a search of Office 365." -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Check for duplicate names in local AD or the CSV file.             " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                $user.ADAccountCreated = "CREATED"
            }
            else
            {
                Try
                {
                    #Read and convert the manager entry into the manager's UserPrincipalName
                    if(($user.Manager -ne $null) -and ($user.Manager -ne ""))
                    {
                        $ManagerDisplayName = '*'+$user.Manager+' | '+$userAccountVariables[1]+'*'
                        $ManagerDN = (Get-ADuser -Filter {displayname -like $ManagerDisplayName}).DistinguishedName
                    }

                    #Create the AD User
                    New-ADUser `
                    -GivenName $user.FirstName `
				    -Surname  $user.LastName `
                    -Initials $user.Initials `
                    -Description $user.AssitantMailAddress `
                    -OfficePhone $userAccountVariables[11] `
                    -MobilePhone $user.MobilePhone `
                    -Fax $userAccountVariables[12] `
                    -EmailAddress $userAccountVariables[7] `
                    -HomePage $userAccountVariables[8] `
                    -StreetAddress $userAccountVariables[9] `
                    -City "CITY" `
                    -PostalCode $userAccountVariables[10] `
                    -Title $user.JobTitle `
                    -Department $user.compartment `
                    -Manager $ManagerDN `
                    -POBox "D" `
                    -State "D" `
                    -Country "DE" `
                    -HomeDirectory $HomeDirectory `
                    -Credential $ADCredential `
                    -Name "$ADObjectName" `
                    -DisplayName $userAccountVariables[3] `
                    -AccountPassword $userAccountVariables[0] `
                    -SAMAccountName $userAccountVariables[4] `
                    -Company $userAccountVariables[1] `
                    -Path $Path `
                    -UserPrincipalName $userAccountVariables[5] #-Verbose -WhatIf

                    #Pause for 5 seconds
                    Start-Sleep 5

                    #Enable the account and allow login without changing password
                    #Enable-ADAccount -Identity $userAccountVariables[4]
                    Set-ADUser -Identity $userAccountVariables[4] -Credential $ADCredential -ChangePasswordAtLogon $false -Enabled $true
                    Set-ADUser -Identity $userAccountVariables[4] -Credential $ADCredential -replace @{ipPhone=$user.DirectPhone}

                    if(($user.NameOfAssistant -ne $null) -or ($user.NameOfAssistant -ne ""))
                    {
                        Set-ADUser -Identity $userAccountVariables[4] -Credential $ADCredential -replace @{info=$user.NameOfAssistant}
                    }

                    $user.ADAccountCreated = "CREATED"
                    $user.ADAccountCreationDate = Get-Date -Format "dd-MM-yyyy"
                }
                Catch
                {
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                    Write-Host "New-ADUser failed to create a local AD User:"$user.FirstName $user.LastName     -ForegroundColor Red -BackgroundColor Black
                    Write-Warning -Message $_.Exception.Message
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                }
            }
            Write-Host "======================================================================================" -ForegroundColor Yellow
        }# End of for loop for the CSV file
        Write-Host "Finished with looping through local AD users in the Onboarding CSV file" -ForegroundColor Cyan
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "General failure setting up the new local AD user account creation steps" -ForegroundColor Red -BackgroundColor Black
        Write-Host "on user:"$user.FirstName $user.LastName                                  -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
    }
    #Write to the CSV to indicate that the account was created in local AD
    $users | Export-CSV -Path $CSVPath -UseCulture -Encoding UTF8 -NoTypeInformation
}

##################################################################
### Create a user in Azure AD using the info from the CSV file ###
##################################################################
function createAzureADUser
{
    param
    (
       [string] $CSVPath
    )

    Write-Host "============================================" -ForegroundColor Yellow
    Write-Host "Provisioning Users in Azure Active Directory" -ForegroundColor Yellow
    Write-Host "============================================" -ForegroundColor Yellow
    
    Try
    {
        #Import the CSV File which is passed as parameter when the script is called - Again, make sure the CSV is ";" delimited and not "," delimited
        $users = Import-CSV -path $CSVPath -Encoding UTF8 -UseCulture

        #Connect to Azure AD using the connectToAzureAD function described above
        connectToAzureAD
    
        #Loop through users in CSV file
        foreach ($user in $users)
        {
			Write-Host "Creating Azure AD user:"$user.FirstName $user.LastName -ForegroundColor Cyan
            #Run the generateUserAccountVariables function to generate an array of needed variables for AD
            $userAccountVariables = generateUserAccountVariables $user.Company $user.FirstName $user.LastName $user.EmailAddress

            #Special password profile variable for storing passwords specifically for Azure AD - It's a feature!
            $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
            $PasswordProfile.Password = $userAccountVariables[0]

            #Check that the user does exist in local AD before creating in Azure AD by comparing SAMAccountNames and the AccountCreated column in the CSV file
            $ADFilter = "userPrincipalName -eq "+''''+$userAccountVariables[5]+''''
            $DoesADUserExist = Get-ADUser -Filter $ADFilter
            if (($user.ADAccountCreated -eq "") -and ($user.ADAccountCreated -ieq "nein") -and ($user.ADAccountCreated -imatch "no"))
            {
                if ($DoesADUserExist -ne $null)
                {
                    #If user does exist in local AD, but is not marked as such in the CSV output a warning message and update CSV
                    Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                    Write-Host "Warning: a local AD user with the UserPrincipalName:               " -ForegroundColor White -BackgroundColor DarkBlue
                    Write-Host  $userAccountVariables[5]" already exists in a search of Office 365." -ForegroundColor White -BackgroundColor DarkBlue
                    Write-Host "Check for duplicate names in Azure AD or the CSV file.             " -ForegroundColor White -BackgroundColor DarkBlue
                    Write-Host "===================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                    $user.ADAccountCreated = "CREATED"
                }
                else
                {
                    #If user does not exist as indicated by the CSV file, output a warning message and continue to the next Azure AD user
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                    Write-Host "Warning: a local AD user with the UserPrincipalName:                         " -ForegroundColor Red -BackgroundColor Black
                    Write-Host  $userAccountVariables[5]" does not exist as indicated by the CSV file.       " -ForegroundColor Red -BackgroundColor Black
                    Write-Host "Skipping this user from Azure AD account creation                            " -ForegroundColor Red -BackgroundColor Black
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                    continue
                }
            }
            elseif ($DoesADUserExist -eq $null)
            {
                #If user does exist in local AD, output a warning message and continue to next user and continue to the next Azure AD user
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                Write-Host "Warning: a local AD user with the UserPrincipalName:                         " -ForegroundColor Red -BackgroundColor Black
                Write-Host  $userAccountVariables[5]" does not exist as indicated by a local AD lookup.  " -ForegroundColor Red -BackgroundColor Black
                Write-Host "Skipping this user from Azure AD account creation                            " -ForegroundColor Red -BackgroundColor Black
                Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                continue
            }

            #Check that the user does not already exist by comparing UserPrincipalNames and the AzureADAccountCreated column in the CSV file
            #But first create a string for the Filter parameter as ObjectID always throws and error and will bork things
            $AzureADFilter = "userPrincipalName eq "+''''+$userAccountVariables[5]+''''
            $DoesAzureADUserExist = Get-AzureADUser -Filter $AzureADFilter
            if (($user.AzureADAccountCreated -ne "") -and ($user.AzureADAccountCreated -ine "nein") -and ($user.AzureADAccountCreated -ine "no"))
            {
                #If user appears to exist as indicated by the CSV file, output a warning message
                Write-Host "Warning: an Azure AD user with the UserPrincipalName:" $userAccountVariables[5] "appears to exist as indicated by the CSV file. Skipping..." -ForegroundColor White -BackgroundColor DarkBlue
            }
            elseif ($DoesAzureADUserExist -ne $null)
            {
                #If user does exist in Office365, output a warning message
                Write-Host "==========================================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: an Azure AD user with the UserPrincipalName: "$userAccountVariables[5]             -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "already exists in a search of Office 365.                                                 " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Check for duplicate names in Azure AD.                                                    " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "The CSV file will be updated accordingly.                                                 " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "==========================================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                $user.AzureADAccountCreated = "CREATED"
            }
            else
            {
                Try
                {
                    #Create the Azure AD User
                    New-AzureADUser `
                    -GivenName $user.FirstName `
				    -Surname  $user.LastName `
                    -TelephoneNumber $userAccountVariables[11] `
                    -Mobile $user.Mobile `
                    -FacsimileTelephoneNumber $userAccountVariables[12] `
                    -StreetAddress $userAccountVariables[9] `
                    -City "CITY" `
                    -PostalCode $userAccountVariables[10] `
                    -Department $user.compartment `
                    -JobTitle $user.JobTitle `
                    -State "D" `
                    -DisplayName $userAccountVariables[3] `
                    -PasswordProfile $PasswordProfile `
                    -UserPrincipalName $userAccountVariables[5] `
                    -AccountEnable $true `
                    -MailNickName $userAccountVariables[6] `
                    -PreferredLanguage "de-DE" `
                    -UsageLocation "DE" `
                    -Country "DE"

                    #-Verbose #-WhatIf

                    Start-Sleep 5

                    #Allow the user to sign-in
                    Set-AzureADUser -ObjectId $userAccountVariables[5] -AccountEnabled $True
                }
                Catch
                {
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                    Write-Host "New-AzureADUser failed to create a local AD User:"$user.FirstName $user.LastName       -ForegroundColor Red -BackgroundColor Black
                    Write-Warning -Message $_.Exception.Message
                    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
                }
            #Set the AccountCreation variables from the CSV to indicate that the account was created in AzureAD
            $user.AzureADAccountCreated = "CREATED"
            $user.AzureADAccountCreationDate = Get-Date -Format "dd-MM-yyyy"
            }
            Write-Host "======================================================================================" -ForegroundColor Yellow
        }#End of for loop for New-AzureADUser
        Write-Host "Finished with looping through Azure AD users in the CSV file" -ForegroundColor Cyan
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "General failure setting up the Azure AD user account creation steps" -ForegroundColor Red -BackgroundColor Black
        Write-Host "on user:"$user.FirstName $user.LastName                              -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
    }
    #Write to the CSV to indicate that the account was created in Azure AD
    $users | Export-CSV -Path $CSVPath -UseCulture -Encoding UTF8 -NoTypeInformation
}

##################################################################
### Assign Office365 licenses using the info from the CSV file ###
##################################################################
function assignOffice365Licenses
{
    param
    (
       [string] $CSVPath
    )

    Try
    {
        Write-Host "==================================" -ForegroundColor Yellow
        Write-Host "Provisioning Licenses in Office365" -ForegroundColor Yellow
        Write-Host "==================================" -ForegroundColor Yellow
        
        #Connect to Azure AD using the connectToAzureAD function described above
        connectToAzureAD

        #Import the CSV File which is passed as parameter when the script is called - Again, make sure the CSV is ";" delimited and not "," delimited
        $users = Import-CSV -path $CSVPath -Encoding UTF8 -UseCulture

        foreach ($user in $users)
        {
			Write-Host "Assigning licenses to Office 365 user:"$user.FirstName $user.LastName -ForegroundColor Cyan
            #Run the generateUserAccountVariables function to generate an array of needed variables for AD
            $userAccountVariables = generateUserAccountVariables $user.Company $user.FirstName $user.LastName $user.EmailAddress

            #Get the standard license type, Office 365 E3
            $officePlanName = "ENTERPRISEPACK"
            $officeLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $officeLicense.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $officePlanName -EQ).SkuID

            $visioPlanName = "VISIOCLIENT"
            $visioLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $visioLicense.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $visioPlanName -EQ).SkuID
        
            #Create a license object array to which the licenses will be added after being checked if they are required/available
            $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        
            #Get the Azure AD Users license info in order to check that they do not already have a license
            $AzureADFilter = "userPrincipalName eq "+''''+$userAccountVariables[5]+''''
            $DoesAzureADUserExist = Get-AzureADUser -Filter $AzureADFilter
            if ($DoesAzureADUserExist -ne $null)
            {
				#Get the users license info to check that they dont already have a license in question
				$LicenseInfo = Get-AzureADUserLicenseDetail -ObjectId  $DoesAzureADUserExist.ObjectId
			}

            #Switch variables to check if licenses need to be added or not
            $AddOfficeLicensesSwitch = $false
            $AddVisioLicensesSwitch = $false

            #Check the CSV to see if the user should receive the standard licenses
            if ($DoesAzureADUserExist -eq $null)
            {
                Write-Host "========================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning:Looks like "$userAccountVariables[5]              -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "is not in Office365 or the account could not be found.  " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "You better check Office 365.                            " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "========================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }
            elseif (($user.StandardOfficeLicense -ne $null) -and ($user.StandardOfficeLicense -ne "") -and ($user.StandardOfficeLicense -ine "nein") -and `
            ($user.StandardOfficeLicense -ine "no") -and ($user.StandardOfficeLicense -ne "ASSIGNED"))
            {
                Write-Host "=============================================" -ForegroundColor Cyan -BackgroundColor DarkBlue
                Write-Host "Warning:Assigning standard office licenses to" -ForegroundColor Cyan -BackgroundColor DarkBlue
                Write-Host $userAccountVariables[5]                        -ForegroundColor Cyan -BackgroundColor DarkBlue
                Write-Host "and updating CSV accordingly.                " -ForegroundColor Cyan -BackgroundColor DarkBlue
                Write-Host "=============================================" -ForegroundColor Cyan -BackgroundColor DarkBlue

                $LicensesToAssign.AddLicenses = $officeLicense
                $AddOfficeLicensesSwitch = $true
            }
            elseif ($LicenseInfo.SkuPartNumber -imatch "ENTERPRISEPACK")
            {
                Write-Host "========================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: Looks like "$userAccountVariables[5]             -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "alread has some licenses assigned.                      " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "You better check Office 365, but here is what they have:" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host $LicenseInfo.SkuPartNumber                                 -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "========================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }
            else
            {
                Write-Host "================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: Looks like "$userAccountVariables[5]     -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "was not marked to receive an Office license.    " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "You better check the CSV file.                  " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "No licenses assigned.                           " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }

            #Check the CSV to see if the user should receive a Visio license
            if ($DoesAzureADUserExist -eq $null)
            {
                Write-Host "========================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: Looks like "$userAccountVariables[5]             -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "is not in Office365 or the account could not be found.  " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "You better check Office 365.                            " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "========================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }
            elseif (($user.VisioLicense -ne $null) -and ($user.VisioLicense -ne "") -and ($user.VisioLicense -ine "nein") -and `
            ($user.VisioLicense -ine "no") -and ($user.VisioLicense -ne "ASSIGNED"))
            {
                Write-Host "==================================" -ForegroundColor Cyan -BackgroundColor DarkBlue
                Write-Host "Assigning a Visio license to      " -ForegroundColor Cyan -BackgroundColor DarkBlue
                Write-Host $userAccountVariables[5]             -ForegroundColor Cyan -BackgroundColor DarkBlue
                Write-Host "and updating CSV accordingly      " -ForegroundColor Cyan -BackgroundColor DarkBlue
                Write-Host "==================================" -ForegroundColor Cyan -BackgroundColor DarkBlue

                $LicensesToAssign.AddLicenses = $visioLicense
                $AddVisioLicensesSwitch = $false
            }
            elseif (($LicenseInfo.SkuPartNumber -imatch "VISIOCLIENT"))
            {
                Write-Host "========================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: Looks like "$userAccountVariables[5]             -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "already has a Visio license assigned.                   " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "You better check Office 365, but here is what they have:" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host $LicenseInfo.SkuPartNumber                                 -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "========================================================" -ForegroundColor White -BackgroundColor DarkBlue
            }
            else
            {
                Write-Host "===============================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: Looks like "$userAccountVariables[5]    -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "was not marked to receive an Visio license.    " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Check the CSV file if this is in error.        " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "===============================================" -ForegroundColor White -BackgroundColor DarkBlue
            }

            #Check if the AddLicensesSwitch variable is switched on and then add the licenses
            if (($AddOfficeLicensesSwitch -eq $true) -and ($AddVisioLicensesSwitch -eq $true))
            {
                Write-Host "Office 365 E3 and Visio licenses were assigned to" $user.FirstName $user.LastName -ForegroundColor Cyan
                #Assign licenses that were checked for above
                Set-AzureADUserLicense -ObjectId $userAccountVariables[5] -AssignedLicenses $LicensesToAssign
                
                $LicensesToAssign = $null
                $AddOfficeLicensesSwitch = $false
                $AddVisioLicensesSwitch = $false

                #Update the CSV to reflect the license assignment
                $user.StandardOfficeLicense = "ASSIGNED"
                $user.VisioLicense = "ASSIGNED"
            }
            elseif (($AddOfficeLicensesSwitch -eq $true) -and ($AddVisioLicensesSwitch -eq $false))
            {
                Write-Host "Office 365 E3 license was assigned to" $user.FirstName $user.LastName -ForegroundColor Cyan
                #Assign licenses that were checked for above
                Set-AzureADUserLicense -ObjectId $userAccountVariables[5] -AssignedLicenses $LicensesToAssign

                $LicensesToAssign = $null
                $AddOfficeLicensesSwitch = $false
                $AddVisioLicensesSwitch = $false

                #Update the CSV to reflect the license assignment
                $user.StandardOfficeLicense = "ASSIGNED"
            }
            elseif (($AddOfficeLicensesSwitch -eq $false) -and ($AddVisioLicensesSwitch -eq $true))
            {
                Write-Host "Visio license was assigned to" $user.FirstName $user.LastName -ForegroundColor Cyan
                #Assign licenses that were checked for above
                Set-AzureADUserLicense -ObjectId $userAccountVariables[5] -AssignedLicenses $LicensesToAssign

                $LicensesToAssign = $null
                $AddOfficeLicensesSwitch = $false
                $AddVisioLicensesSwitch = $false

                #Update the CSV to reflect the license assignment
                $user.VisioLicense = "ASSIGNED"
            }
            else
            {
                Write-Host "No licenses were assigned to" $user.FirstName $user.LastName -ForegroundColor Cyan
            }
            Write-Host "======================================================================================" -ForegroundColor Yellow
        }#End of for loop for Set-AzureADUserLicense
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "Error assigning Office365 licenses to the user" $userAccountVariables[5]      -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
    }
    $users | Export-CSV -Path $CSVPath -UseCulture -Encoding UTF8 -NoTypeInformation
}

####################################################
### Enable the Exchange Online archiving feature ###
####################################################
function enableExchangeOnlineArchive
{
    param
    (
       [string] $CSVPath
    )

    Try
    {
        Write-Host "=====================================" -ForegroundColor Yellow
        Write-Host "Enabling Exchange Online Mail Archive" -ForegroundColor Yellow
        Write-Host "=====================================" -ForegroundColor Yellow
        
        #Connect to Exchange Online
        $Session = connectToExchangeOnline
        Import-PSSession $Session -DisableNameChecking 

        #Import the CSV File which is passed as parameter when the script is called - Again, make sure the CSV is ";" delimited and not "," delimited
        $users = Import-CSV -path $CSVPath -Encoding UTF8 -UseCulture

        #Loop through users in CSV file
        foreach ($user in $users)
        {
			Write-Host "Checking Exchange Online archiving for:"$user.FirstName $user.LastName -ForegroundColor Cyan
            
            #Run the generateUserAccountVariables function to generate an array of needed variables for Exchange Online
            $userAccountVariables = generateUserAccountVariables $user.Company $user.FirstName $user.LastName $user.EmailAddress

            #Check that the user's mailbox does exist and if not try again after 10 minutes
            $ExchangeMailbox = Get-Mailbox -Identity $userAccountVariables[5] -ErrorAction SilentlyContinue

            #Flag variable to confirm if the user's mailbox exists
            $UsersMailboxExists = $false

            if ($ExchangeMailbox -eq $null)
            {
                #If the user's mailbox is not yet there or doesn't exist, output a warning message
                Write-Host "===============================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Warning: there is no mailbox for UserPrincipalName:" $userAccountVariables[5]    -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Check for typos or misspelled names in Azure AD or the CSV file.               " -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "Waiting for  10 minutes before trying again as the mailbox may not yet be ready" -ForegroundColor White -BackgroundColor DarkBlue
                Write-Host "===============================================================================" -ForegroundColor White -BackgroundColor DarkBlue

                #Go to sleep for 10 minutes (600 seconds)
                Start-Sleep -s 600

                #Try again to get the user's mailbox
                $ExchangeMailbox = Get-Mailbox -Identity $userAccountVariables[5] -ErrorAction SilentlyContinue
                if ($ExchangeMailbox -eq $null)
                {
                    #If after 10 minutes of waiting the user's mailbox is still not there, output a warning message and break
                    Write-Host "=========================================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                    Write-Host "Warning: even after waiting for 10 minutes there is no mailbox for UserPrincipalName:"     -ForegroundColor White -BackgroundColor DarkBlue
                    Write-Host $userAccountVariables[5] "Check for typos or misspelled names in Azure AD or the CSV file." -ForegroundColor White -BackgroundColor DarkBlue
                    Write-Host "Skipping...                                                                              " -ForegroundColor White -BackgroundColor DarkBlue
                    Write-Host "=========================================================================================" -ForegroundColor White -BackgroundColor DarkBlue
                    break
                }
                else
                {
                    $UsersMailboxExists = $true
                }
            }
            else
            {
                $UsersMailboxExists = $true
            }

            #If the user's mailbox does exist, turn on the Exchange Online archiving feature
            if(($UsersMailboxExists -eq $true) -and ($ExchangeMailbox.ArchiveStatus -eq "None"))
            {
                Write-Host "Enabling Exchange Online archiving for:"$ExchangeMailbox.UserPrincipalName -ForegroundColor Cyan
                Enable-Mailbox -Identity $userAccountVariables[5] -Archive | Out-Null
            }
            elseif(($UsersMailboxExists -eq $true) -and ($ExchangeMailbox.ArchiveStatus -eq "Active"))
            {
                Write-Host "Exchange Online archiving was already enabled for:"$ExchangeMailbox.UserPrincipalName "Skipping..." -ForegroundColor Cyan
            }
        }
    }
    Catch
    {
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black
        Write-Host "General Error enalbing Exchange Online archiving feature for " $userAccountVariables[5]  -ForegroundColor Red -BackgroundColor Black
        Write-Warning -Message $_.Exception.Message
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red -BackgroundColor Black

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
$LogPath = "\\DOMAIN.NAME\Admins\Scripts\On- and Offboarding\Logs\Onboarding\"
$Daysback = "-90"
 
$CurrentDate = Get-Date
$DatetoDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $LogPath | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
}

#########################################################################
### GROUP MEMBERSHIP ?????????????????????????????????????????????????###
#########################################################################

#########################################################################
### PROXY ADDRESSES ??????????????????????????????????????????????????###
#########################################################################

#########################################################################
### Disconnect from open services and close running transcription log ###
#########################################################################
function endLogging
{
    
    #Test if there is a connection to Azure AD and if so disconnect
    Try
    {
        $checkifConnectedToAzureAD = Get-AzureADTenantDetail -ErrorAction SilentlyContinue

        #Discconect from AAD
        Disconnect-AzureAD
        Write-Host "Disconnected from Azure AD" -ForegroundColor Cyan
    }
    Catch
    {
        Write-Host "Already disconnected from Azure AD" -ForegroundColor Cyan
    }

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

#Create a local AD user
createLocalADUser($CSVPath)

#Create an Azure AD user
createAzureADUser($CSVPath)

#Assign said Azure AD user Office and Visio licenses depending on their CSV file data
assignOffice365Licenses($CSVPath)

#Turn on the Exchange Online archiving feature
enableExchangeOnlineArchive($CSVPath)

#Clean up log files older than 90 days
deleteOldLogs

#Stop the transcript and save
endLogging