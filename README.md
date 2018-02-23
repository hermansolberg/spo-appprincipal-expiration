
# SharePoint Online app principal expiration monitoring
This repo contains a PowerShell script meant to be run as a PowerShell Runbook in Azure Automation. It will check, based on configured variables, if there are any Provider-Hosted Add-ins whose app principals have expired, or will expire within the next 100 days.

## Dependencies
The Runbook required the following dependencies to be configured:
* An Azure Automation Account with a Run As account created
* The Service Principal for the Run As account need the "WindowAzureActiveDirectory/Application/Read directory data" permission granted
* The AzureAD PowerShell module from the modules gallery imported into the Automation Account
