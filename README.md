# ExoHelper Powershell module
Simple wrapper module that calls into EXO REST API without the need for full heavy-weight ExchangeOnlineManagent module.

# Usage
Usage of module is pretty simple:
- create authentication factory with clientId that has proper permissions granted (for app-only, Exchange.ManageAsApp permission)
- call Initialize-ExoAuthentication command to create a connection to EXO
- call Invoke-ExoCommand, passing name of command, and hashtable with command parameters

_Note_: Module relies on [AadAuthenticationFactory](https://github.com/GreyCorbel/AadAuthenticationFactory) module that implements necessary authentication flows for AAD.

Sample below:
```powershell
#create authentication factory
$appId = 'xxx' #app id of app registration that has appropriate permissions granted for EXO app-only management
$clientSecret = 'yyy'   #client secret for app registration
$tenantId = 'mydomain.onmicrosoft.com'

$factory = New-AadAuthenticationFactory -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
#initialize the ExoHelper module
$Connection = New-ExoConnection -Authenticationfactory $factory -TenantId $tenantId

#call EXO command
$params = @{
    Identity = "myuser@mydomain.com"
}
#Specification of connection is optional here - module automatically uses last connection created when explicit connection not provided
Invoke-ExoCommand -Name 'Get-Mailbox' -Parameters $params -Connection $Connection

```
