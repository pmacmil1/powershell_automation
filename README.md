# powershell_automation
Scripts used in the automation various Office365 and Azure AD tasks.  Note: most of these scripts use an encrypted credential .key file for storing the least-priviledged admin account that can manipulate either Azure AD or Sharepoint.  Please check [here](https://www.pdq.com/blog/secure-password-with-powershell-encrypting-credentials-part-1/) for information on how to generate these encrypted files and how to restrict access to them.

## Onboarding and Offboarding of Users
The Onboarding and Offboarding scripts are to be used for complex company structures where users for multiple subsidiary companies are managed by one central IT department.  Users parameters is to be entered into a centralized CSV file by HR, which is then read by the script to populate the various parameters needed by the user account creation process.

The onboarding script will create local AD users and Azure AD users for synchronization with Office365.  Further the script will assign group membership and Office365 licenses based on the input read from the CSV file.

The Offboarding script will rollback the user creation process, devprovisioning the user from all licenses and groups, while providing manager access to their Exchange and Sharepoint Online accounts.

## Remove Deactivated Users From O365 Groups
This script checks for deactivated users in Azure AD and then removes them for Exchange Distribution Groups, Office365 Groups, and Shared Mailboxes.

## Sharepoint Online (SPO) File Uploads
This script will take a locally hosted folder and recursively upload the contents to the Sharepoint Online library of your choosing.  It will create the root and subfolders of the folder in question, and create them while uploading their file contents.  In this case it will only upload PDF file types but can be configured for any other file types.

## Sharepoint Online (SPO) User Info Updates
Because Azure AD and the SPO userstore are linked but ultimately separate, it can be difficult to update some fields in SPO which are not synced with Azure AD.

You can read more about this [here](https://blog.atwork.at/post/SharePoint-Online-UserProfiles-and-the-story-about-synchronizing-with-Azure-Active-Directory).

This script will directly update the user fields in SPO that are not available or synced from Azure AD for SPO.  Here it is updated telephone numbers and SIP addresses for a SPO-hosted user directory.
