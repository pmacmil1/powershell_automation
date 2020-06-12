# powershell_automation
Scripts used in the automation various Office365 and Azure AD tasks.

## Onboarding and Offboarding of Users

The Onboarding and Offboarding scripts are to be used for complex company structures where users for multiple subsidiary companies are managed by one central IT department.  Users parameters is to be entered into a centralized CSV file by HR, which is then read by the script to populate the various parameters needed by the user account creation process.

The onboarding script will create local AD users and Azure AD users for synchronization with Office365.  Further the script will assign group membership and Office365 licenses based on the input read from the CSV file.

The Offboarding script will rollback the user creation process, devprovisioning the user from all licenses and groups, while providing manager access to their Exchange and Sharepoint Online accounts.

## Sharepoint File Uploader


## Sharepoint Online (SPO) User Info Updates
Because Azure AD and the SPO userstore are linked but ultimately separate, it can be difficult to update some fields in SPO which are not synced with Azure AD.

You can read more about this [here](https://blog.atwork.at/post/SharePoint-Online-UserProfiles-and-the-story-about-synchronizing-with-Azure-Active-Directory).

This script will directly update the user fields in SPO that are not available or synced from Azure AD for SPO, 
