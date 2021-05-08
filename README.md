# Graph API with PowerShell

As System Lead for SharePoint Online in the the company I’m working get very often the question how to get the access token to use Microsoft Graph API. I’m more familiar with the program language PowerShell I will do a short tutorial with an example script how I’m connecting to Microsoft Graph API in my scripts I’m using as System Lead.

Please be aware there are two different types of permission you can have for your REST API endpoints in Graph API. You will find **application permission** and **delegated permission**. With application permission you have full access to all e.g. sites on your *Microsoft 365* tenant without any user context. With delegated permission you have only this permission setup in your Azure Application you have with the logged in user. That’s the reason you need to get the access token for delegated permission a user in your request.

## Example

You will find an example of the following explanation in the Git repository called `Connect-PStoGraphAPIDelegated.ps1`.

## 1. Prepare the body for the request

First you have to prepare the body before you start requesting the access token. The body have to look like this if the user context is needed:

``` PowerShell
$reqTokenBody = @{
    Grant_Type    = "Password"
    client_Id     = $ClientId
    Client_Secret = $secret.GetNetworkCredential().Password
    Username      = $UserName
    Password      = $password.GetNetworkCredential().Password
    Scope         = "https://graph.microsoft.com/.default"
}
```

The *Grant_Type* is always from the value `"Password"` and you need the clients id and secret of the registered Azure application and the username and password for your context to do your action on the SharePoint Online sites. The user must be a member or owner of this site on the tenant. The *Scope* is again defined with `"https://graph.microsoft.com/.default"`.

## 2. Send the request to Microsoft

To get the access token from Microsoft you have to send the created body to Microsoft. In PowerShell you are doing this with the following command:

``` PowerShell
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method POST -Body $reqTokenBody
```

Uri of this request depends of your tenant name.

## 3. Prepare the header for your requests

If you was getting your response from Microsoft with your access token you have to prepare your header for your requests.

``` PowerShell
$graphApiHeader = @{ Authorization = "Bearer $tokenResponse.access_token" }
```

Important: The access token is valid for one hour. When the time is over you have to request for a new access token!

## Done

Now you can start with your request to the Graph API endpoints.