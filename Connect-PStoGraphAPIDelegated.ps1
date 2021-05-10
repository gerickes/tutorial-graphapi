[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)][String]$Tenant
)
<#
  .SYNOPSIS
    Example to connect to Graph API.

  .DESCRIPTION
    This PowerShell script is an example of how to connect to Microsoft Graph API with PowerShell.
    User context is needed because of delegation permission are defined in the Azure Application.

  .COMPONENT
    Windows PowerShell 5.1 or PowerShell Core

  .PARAMETER Tenant
    This parameter is mandatory must include the name of the tenant which is this:
    [https://<Tenant>.sharepoint.com/]

  .NOTES
    Version:          1.0
    Author:           Stefan Gericke - stefan@gericke.cloud
    Creation Date:    2021/05/08
    Description:      Creating of the example script

  .EXAMPLE
    Connect-PStoGraphAPIDelegated -Tenant <name of the tenant>

#>

#------------------------------------------------[Function]-------------------------------------------------------

#----------------------------------------------[Declarations]-----------------------------------------------------

# IMPORTANT: Needed for the Federation Server of our company which speaks only TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Constant
$graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

#------------------------------------------------[Execution]------------------------------------------------------

Write-Host "*** Start PowerShell script Connect-PStoGraphAPIDelegated.ps1 ***"

# Ask for the password of the user and the secretof the regeistered Azure application.
$credUser = Get-Credential -Message "Enter credentials for the user context"
$credAzureAplication = Get-Credential -Message "Enter the client id and client secret of the Azure application"

# Prepare the body before request for the access token
# UserName and password is needed for the user context. At the end you will get access only to SharePoint Online Sites
# where the user is added as member or owner of the site.
$tenantName = "$Tenant.onmicrosoft.com"
$reqTokenBody = @{
    Grant_Type    = "Password"
    client_Id     = $credAzureAplication.UserName
    Client_Secret = $credAzureAplication.GetNetworkCredential().Password
    Username      = $credUser.UserName
    Password      = $credUser.GetNetworkCredential().Password
    Scope         = "https://graph.microsoft.com/.default"
}
Write-Host "The body is created for the getting the access token ..."

# Start request for getting the access token from Graph API
try {
    Write-Host "Getting token for Microsoft Graph Teams App ..."
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method POST -Body $reqTokenBody
    Write-Host "The token has been received successful!"
}
catch {
    Write-Error "Error on trying to connect to Graph API!"
    Write-Error $_.Exception.Message
    throw Exception
} # try .. catch: Start request for getting the access token from Graph API

# If access token is available continue with your request to your endpoints
if ($tokenResponse.access_token) { # access token is available

    # Create header for using Graph API
    $graphApiHeader = @{ Authorization = "Bearer $tokenResponse.access_token" }

    # Here an example for a request to an endpoint.
    # Get the id of the user in Azure AD
    $getUserIdUri = "$($graphApiBaseUrl)/users/$($credUser.UserName)"
    Write-Host "Send web request: $getUserIdUri ..."
    try {
        $webRequest = Invoke-RestMethod -Headers $graphApiHeader -Uri $getUserIdUri -Method Get -ContentType "application/json"
        $userId = $webRequest.id
        Write-Host "User Id of $Username in Azure AD: $userId"
    } catch {
        Write-Error "The request to Graph API wasn´s successful!"
        Write-Error $_.Exception.Message
    } # try .. catch: Get the id of the user in Azure AD
} # if: access token is available

Write-Host "*** Stop PowerShell script Connect-PStoGraphAPIDelegated.ps1 ***"
