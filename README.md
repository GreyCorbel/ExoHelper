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
#initialize the Exo connection
$Connection = New-ExoConnection -Authenticationfactory $factory -TenantId $tenantId

#call EXO command
$params = @{
    Identity = "myuser@mydomain.com"
}
#Specification of connection is optional here
#Module automatically uses last connection created when explicit connection not provided
Invoke-ExoCommand -Name 'Get-Mailbox' -Parameters $params -Connection $Connection

```
_Note_: To protect sensitive data (e.g. passwords to be set on newly created mailboxes), Exchange Online uses RSA Key pair with public key embedded into temporary module that dynamically downloads when running `Connect-ExchangeOnline`:
![image](https://github.com/user-attachments/assets/492d9293-1d1a-4500-9d49-2c96f73a264a)
Key pair is occassionally rotated. To allow usage of commands that work with sentitive information in ExoHelper module, public key that comes with Exchange Online module is also stored and regularly refreshed in this repo, and ExoHelper module loads it from here when imported, and caches on it on machine where it is executed. When not able to download the public key, or cached key gets outdated, module still works, but commands that work with sensitive data will fail.
I wish Microsoft would allow retrieval of publis key directly from their REST API!
