# Parse required command line arguments: Welkin client id/secret AD app id and tenant id, AD app client id/secret, consumption location [defaults to westus]
Param (
[Parameter(Mandatory=$true,HelpMessage="The client ID generated for your Welkin company")][string]$WelkinClientId,
[Parameter(Mandatory=$true,HelpMessage="The client secret generated for your Welkin company")][string]$WelkinClientSecret,
[Parameter(Mandatory=$true,HelpMessage="The ID of your M365 tenant")][string]$TenantId,
[Parameter(Mandatory=$true,HelpMessage="The client ID generated for your M365 client application")][string]$AppClientId,
[Parameter(Mandatory=$true,HelpMessage="The client secret generated for M365 client application")][string]$AppClientSecret,
[Parameter(HelpMessage="The Azure consumption region in which to create resources (defaults to westus)")][string]$Region = "westus"
)

Write-Host $WelkinClientId
Write-Host $WelkinClientSecret
Write-Host $TenantId
Write-Host $AppClientId
Write-Host $AppClientSecret
Write-Host $Region

# Log in using an Azure admin account for the subscription you wish to deploy to
az login

# Create a resource group named OutlookWelkinSyncResourceGroup in the given consumption location if none exists
$ResourceGroup = "OutlookWelkinSyncResourceGroup"
$Exists = az group exists -n $ResourceGroup

if ($Exists -eq $false) {
    Write-Host "Creating resource group $ResourceGroup"
    az group create -n $ResourceGroup -l $Region
}

# Create a key vault named SyncKeyVault in OutlookWelkinSyncResourceGroup if none exists
$KeyVault = "SyncKeyVault"
$Json =  (az keyvault show -n $KeyVault) | Out-String
$Json = [string]::join("",($Json.Split("`n"))).Trim()
$Exists = $Json.StartsWith('{')

if ($Exists -eq $false) {
    Write-Host "Creating key vault $KeyVault"
    az keyvault create -n $KeyVault -g $ResourceGroup
}

# Add all client creds and app info to the vault with the expected value names, or update if they already exist
Write-Host "Adding client credentials and app info to key vault..."
az keyvault secret set -n "WelkinClientId" --vault-name $KeyVault --value $WelkinClientId
az keyvault secret set -n "WelkinSecretId" --vault-name $KeyVault --value $WelkinSecretId
az keyvault secret set -n "AppClientId" --vault-name $KeyVault --value $AppClientId
az keyvault secret set -n "AppSecretId" --vault-name $KeyVault --value $AppSecretId
az keyvault secret set -n "TenantId" --vault-name $KeyVault --value $TenantId
Write-Host "...credentials added."

# Create a storage account named SyncStorage in OutlookWelkinSyncResourceGroup if none already exists

# Create app insights SyncInsights in OutlookWelkinSyncResourceGroup if it doesn't already exist

# Create the function OutlookWelkinSyncFunction in OutlookWelkinSyncResourceGroup if it doesn't already exist

# Ensure that the function has a managed identity (may need to create an AD tenant first)

# Grant the managed identity read access to SyncKeyVault

# Deploy current folder to OutlookWelkinSyncFunction