<#
  .SYNOPSIS
  This script will create and run a Content Search, then delete the matching emails. 
   
  Adrian Dolder - 365cloud (365cloud.pro)
  	
  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
  Version 1.0, 2025-01-16

  Ideas, comments and suggestions to blog@doudi.ch including the developer name. 
 
  .LINK  
  OneDrive Doudisblog
	
  .DESCRIPTION
  
  This script that allows you to input the sender's email address and the subject
  of the emails you want to delete. This script will create and run a Content Search,
  then delete the matching emails.
  
  This script is meant to be run once.
    
  .NOTES 
  Requirements 
  - PowerShell 5+
  - eDiscovery Manager or Compliance Search role to create and run a Content Search
  - Organization Management (Compliance) or Search And Purge role to delete messages.
  
  Instructions
  - Run Script in PowerShell
  - Enter Credentials of your Administrator with the necessary roles
  - Enter sender address and subject for the email to delete

  Revision History 
  -------------------------------------------------------------------------------- 
  1.0     Initial community release
	
  .PARAMETER
  -
   
  .EXAMPLE
  .\Delete-SpecificEmailInAllMailboxes.ps1

  #>


# Check if the ExchangeOnlineManagement module is installed
$moduleName = "ExchangeOnlineManagement"
$module = Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue

if ($module) {
    Write-Host "$moduleName module is installed. Checking for updates..."
    # Check for updates
    $update = Update-Module -Name $moduleName -ErrorAction SilentlyContinue
    if ($update) {
        Write-Host "$moduleName module has been updated."
    } else {
        Write-Host "$moduleName module is already up to date."
    }
} else {
    Write-Host "$moduleName module is not installed. Installing..."
    # Install the module
    Install-Module -Name $moduleName -Scope CurrentUser -Force
    Write-Host "$moduleName module has been installed."
}


# Connect to Security & Compliance Center
Connect-IPPSSession

# Create Compliance Search
$SenderAddress = Read-Host "E-Mail Adresse des Absenders angeben!"
$emailSubject = Read-Host "Betreff des zu löschenden E-Mails angeben!"

# Create a unique search name
$searchName = "Search_" + (Get-Date).ToString("yyyyMMdd_HHmmss")

# Create the Content Search
New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery "from:$senderAddress AND subject:'$emailSubject'"

# Start the Content Search
Start-ComplianceSearch -Identity $searchName

# Wait for the search to complete
$searchStatus = Get-ComplianceSearch -Identity $searchName
while ($searchStatus.Status -ne "Completed") {
    Write-Host "Waiting for search to complete..."
    Start-Sleep -Seconds 30
    $searchStatus = Get-ComplianceSearch -Identity $searchName
}

# Delete the emails
New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete

Write-Host "Emails from $senderAddress with subject '$emailSubject' have been deleted."