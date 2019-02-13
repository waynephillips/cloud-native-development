# Client secret is the app registation secret in azure ad
Param(
	[string] $ResourceGroupName = $env:ResourceGroupName,
	[string] $AppServiceName = $env:AppServiceName,
    	[string] $AadTenantId = $env:TenantId,
    	[string] $RegistrationSecret = $env:ClientSecret,
    	[string] $RegistrationId = $env:RegistrationId 
)

if(!$ResourceGroupName)
{
	throw "ResourceGroupName is null or empty."
}
if(!$AppServiceName)
{
	throw "AppServiceName is null or empty."
}
if(!$AadTenantId)
{
	throw "AadTenantId is null or empty."
}
if(!$RegistrationSecret)
{
	throw "RegistrationSecret is null or empty."
}
if(!$RegistrationId)
{
	throw "RegistrationId is null or empty."
}



Write-Host "Configuring App Service Authentication for $AppServiceName..."

$AppService = Get-AzureRmWebApp -Name $AppServiceName
$PrimaryWebAppHostName = $($AppService.HostNames[0])
# Get current configuration
$AppAuthSettingsResource = Invoke-AzureRmResourceAction `
		-ResourceGroupName $ResourceGroupName `
		-ResourceType Microsoft.Web/sites/config `
		-ResourceName "$($AppService.Name)/authsettings" `
		-Action list `
		-ApiVersion 2016-08-01 `
        	-Force

$PropertiesObject = $AppAuthSettingsResource.Properties
# Adjust properties
   
$PropertiesObject.enabled = $true
$PropertiesObject.unauthenticatedClientAction = "RedirectToLoginPage"
$PropertiesObject.tokenStoreEnabled = $true

# Allowed redirect urls, but only required for "other" domain names
$PropertiesObject.allowedExternalRedirectUrls = @( 
    "https://$PrimaryWebAppHostName/.auth/login/aad/callback/*"
)


$PropertiesObject.defaultProvider = "AzureActiveDirectory"
$PropertiesObject.clientId = $RegistrationId
$PropertiesObject.clientSecret = $RegistrationSecret  
$PropertiesObject.issuer = "https://sts.windows.net/$AadTenantId/"
$PropertiesObject.isAadAutoProvisioned = $false
$PropertiesObject.additionalLoginParams = @( "response_type=code id_token" ) 

# Extend the refresh token period (optional and depends on data senstivity)
$tokenRefreshExtensionHours = 360 # Note: Default value is 72 hours
if($PropertiesObject.tokenRefreshExtensionHours) {
    $PropertiesObject.tokenRefreshExtensionHours = $tokenRefreshExtensionHours 
}
else {
    $PropertiesObject | Add-Member -Name 'tokenRefreshExtensionHours' -Type NoteProperty -Value $tokenRefreshExtensionHours
}

# Save new properties back to the App Service
New-AzureRmResource -PropertyObject $PropertiesObject `
					-ResourceGroupName $ResourceGroupName `
					-ResourceType Microsoft.Web/sites/config `
					-ResourceName "$($AppService.Name)/authsettings" `
					-ApiVersion 2016-08-01 -Force | Out-Null
