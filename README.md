# ExoHelper Powershell module
Simple wrapper module that calls into EXO REST API without the need for full heavy-weight ExchangeOnlineManagent module.
# Motivation
This module was created to address pain points we observed in ExchangeOnlineManagement module provided by Microsoft:
- download of heavy part of module every time when connecting to EXO - this was not practical for automation tasks that require quick response
- insuficient handling of retry responses from EXO REST API (at that time; might get better meanwhile)
- insuficient handling of EXO REST API non-standard responses (sometimes returns non-JSON response that's not correctly handled)

Design goal was to provide lightweight module that loads quickly, is focused on talking with EXO REST API, provides better error handling and returns as much as specific error information when error occurs.  
Additional benefit is ability to limit fields returned by API via `-PropertiesToLoad` parameter, to limit bandwidth consumption between client and REST API.

Module basically ofers single command `Invoke-ExoCommand` that takes 2 parameters:
- Command name
- Command parameters. Names and types of parameters are the same as accepted by ExchangeOnlineManagement module. Parameters are passed as hashtable.

This generic approach allows implementation of most ExcgangeOnlineManagement commands via this single command.

# Usage
Usage of module is pretty simple:
- create authentication factory with clientId that has proper permissions granted (for app-only, Exchange.ManageAsApp permission)
- call Initialize-ExoAuthentication command to create a connection to EXO
- call Invoke-ExoCommand, passing name of command, and hashtable with command parameters

_Note_: Module relies on [AadAuthenticationFactory](https://github.com/GreyCorbel/AadAuthenticationFactory) module that implements necessary authentication flows for AAD.

# Samples
## App-Only context for backend automation
```powershell
#create authentication factory
$appId = 'xxx' #app id of app registration that has appropriate permissions granted for EXO app-only management
$clientSecret = 'yyy'   #client secret for app registration
$tenantId = 'mydomain.onmicrosoft.com'

$factory = New-AadAuthenticationFactory -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
#initialize the Exo connection. Tenant ID is taken from instance of AAD AuthenticationFactory when not specified explicitly
$Connection = New-ExoConnection -Authenticationfactory $factory

#call EXO command
$params = @{
    Identity = "myuser@mydomain.com"
}
#Specification of connection is optional here
#Module automatically uses last connection created when explicit connection not provided
Invoke-ExoCommand -Name 'Get-Mailbox' -Parameters $params -Connection $Connection
```
## User context for ad-hoc querying
```powershell
$tenantId = 'mydomain.onmicrosoft.com'
New-AadAuthenticationFactory -TenantId $tenantId -ClientId (Get-ExoDefaultClientId) -AuthMode WAM -Name exo
New-ExoConnection -Authenticationfactory exo

Invoke-ExoCommand -Name 'Set-Mailbox' -Parameters {Identity = 'myuser@domain.com'; RetentionPolicy = 'MyRetentionPolicy}
```
# Notes

_Note_: To protect sensitive data (e.g. passwords to be set on newly created mailboxes), Exchange Online uses RSA Key pair with public key embedded into temporary module that dynamically downloads when running `Connect-ExchangeOnline`:
![image](https://github.com/user-attachments/assets/492d9293-1d1a-4500-9d49-2c96f73a264a)
Key pair is occassionally rotated. To allow usage of commands that work with sentitive information in ExoHelper module, public key that comes with Exchange Online module is also stored and regularly refreshed in this repo, and ExoHelper module loads it from here when imported, and caches on it on machine where it is executed. When not able to download the public key, or cached key gets outdated, module still works, but commands that work with sensitive data will fail.
I wish Microsoft would allow retrieval of publis key directly from their REST API!
