function Initialize-ExoAuth
{
<#
.SYNOPSIS
    Initializes EXO connection

.DESCRIPTION
    Initializes EXO connection

.OUTPUTS
    None

.EXAMPLE
Initialize-ExoAuth -authenticationfactory $factory -TenantId mydomain.onmicrosoft.com

Description
-----------
This command initializes connection to EXO REST API.
It uses instance of AADAuthenticationFactory for authentication with EXO REST API
#>
param
    (
        [Parameter(Mandatory)]
        #AAD authentication factory created via New-AadAuthenticationFactory
        #for user context, user factory created with clientId = fb78d390-0c51-40cd-8e17-fdbfab77341b (clientId of ExchangeOnlineManagement module) or your app 
        $AuthenticationFactory,
        
        [Parameter(Mandatory)]
        #Tenant ID - tenant native domain (xxx.onmicrosoft.com)
        [string]
        $TenantId,
        
        [Parameter()]
        #UPN of anchor mailbox
        #Default: UPN of caller or static system mailbox  (for app-only context)
        [string]
        $AnchorMailbox
    )

    process
    {
        $script:ConnectionContext = [PSCustomObject]@{
            AuthenticationFactory = $AuthenticationFactory
            TenantId = $tenantId
            AnchorMailbox = "UPN:$anchorMailbox"
            ConnectionId = [Guid]::NewGuid().ToString()
            ConnectionUri = "https://outlook.office365.com/adminapi/beta/$tenantId/InvokeCommand"
        }

        if([string]::IsNullOrEmpty($AnchorMailbox))
        {
            $claims = Get-ExoToken | Test-AadToken -PayloadOnly
            if($null -ne $claims.upn)
            {
                #using caller's mailbox
                $anchorMailbox = $claims.upn
            }
            else
            {
                #likely app-only context - use same static anchor mailbox as ExchangeOnlineManagement module uses
                $anchorMailbox = "SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}@$tenantId"
            }
        }
        $script:ConnectionContext.AnchorMailbox = "UPN:$anchorMailbox"
    }
}

function Get-ExoToken
{
<#
.SYNOPSIS
    Retrieves access token for authentication with EXO REST API

.DESCRIPTION
    Retrieves access token for authentication with EXO REST API via authentication factory

.OUTPUTS
    Hash table with authorization header containing access token, ready to be passed as headers to web request

.EXAMPLE
Get-ExoToken

Description
-----------
Retieve authorizatin header for calling EXO REST API
#>
param
    (
    )

    process
    {
        Get-AadToken -Factory $script:ConnectionContext.AuthenticationFactory -Scopes 'https://outlook.office365.com/.default' -AsHashTable
    }
}

function Invoke-ExoCommand
{
<#
.SYNOPSIS
    Invokes EXO REST API to execute command provided

.DESCRIPTION
    Invokes EXO REST API to execute command provided along with parameters for the command and optional list of properties to return if full object is not desired

.OUTPUTS
    Data returned by executed command

.EXAMPLE
Invoke-ExoCommand -Name 'Get-Mailbox' -Parameters @{Identity = 'JohnDoe'} -PropertiesToLoad 'netId'

Description
-----------
This command retrieves mailbox of user JohnDoe and returns just netId property
#>
    param
    (
        [Parameter(Mandatory)]
        [string]
            #Name of the command to execute
        $Name,
        
        [Parameter()]
        [hashtable]
            #Hashtable with parameters of the command
        $Parameters = @{},

        [Parameter()]
        [string[]]
            #List of properties to return if not interested in full object
        $PropertiesToLoad,

        [Parameter()]
        [int]
            #Max retries when throttling occurs
        $MaxRetries = 10,

        [switch]
            #If we want to write any warnings returned by EXO REST API
        $ShowWarnings
    )

    begin
    {
        $headers = Get-ExoToken
        $body = @{}
        $batchSize = 1000
        $uri = $script:ConnectionContext.ConnectionUri
        if($PropertiesToLoad.Count -gt 0)
        {
            $props = $PropertiesToLoad -join ','
            $uri = "$uri`?`$select=$props"
        }
        #do not show progress from Invoke-WebRequest
        $pref = $progressPreference
        $progressPreference = 'SilentlyContinue'
    }

    process
    {
        $headers['X-CmdletName'] = $Name
        $headers['client-request-id'] = [Guid]::NewGuid().ToString()
        $headers['Prefer'] = "odata.maxpagesize=$batchSize"
        $headers['connection-id'] = $script:ConnectionContext.connectionId
        $headers['X-AnchorMailbox'] =$script:ConnectionContext.anchorMailbox

        #make sure that hashTable in parameters is properly decorated
        foreach($key in $Parameters.Keys)
        {
            if($Parameters[$key] -is [hashtable])
            {
                $Parameters[$key]['@odata.type'] =  '#Exchange.GenericHashTable'
            }
        }
        $body['CmdletInput'] = @{
            CmdletName = $Name
            Parameters = $Parameters
        }
        $retries = 0
        do
        {
            try {
                $response = Invoke-WebRequest -Uri $uri -Method Post -Body ($body | ConvertTo-Json -Depth 9) -Headers $headers -ContentType 'application/json' -ErrorAction Stop
                #we may process the headers in the future to see rate limit remaining, etc.
                $headers = $response.Headers

                $responseData = $response.Content | ConvertFrom-Json
                
                if($ShowWarnings)
                {
                    foreach($warning in $responseData.'@adminapi.warnings')
                    {
                        Write-Warning $warning
                    }
                }
                $responseData.value
                break
            }
            catch  {
                $ex = $_.exception
                if($PSVersionTable.psEdition -eq 'Desktop')
                {
                    if($ex -is [System.Net.WebException])
                    {
                        if($ex.response.statusCode -ne 429)
                        {
                            throw
                        }
                    }
                    else
                    {
                        #different exception
                        throw
                    }
                }
                else
                {
                    #Core
                    if($ex -is [Microsoft.PowerShell.Commands.HttpResponseException])
                    {
                        if($ex.statusCode -ne 429)
                        {
                            throw
                        }
                    }
                    else
                    {
                        #different exception type
                        throw
                    }
                }
                $headers = $ex.Response.Headers
                $retries++
                Write-Verbose "Retry #$retries"
                if($null -ne $headers['Rate-Limit-Remaining'] -and $null -ne $headers['Rate-Limit-Reset'])
                {
                    if($PSVersionTable.psEdition -eq 'Desktop')
                    {
                        Write-Verbose "Rate limit remaining: $($headers['Rate-Limit-Remaining'])`tRate limit reset: $($headers['Rate-Limit-Reset'])"
                    }
                    else
                    {
                        #Core
                        Write-Verbose "Rate limit remaining: $($headers['Rate-Limit-Remaining'][0])`tRate limit reset: $($headers['Rate-Limit-Reset'][0])"
                    }
                }
                if($retries -gt $MaxRetries)
                {
                    #max retries exhausted
                    throw
                }
                #wait some time
                Start-Sleep -Seconds $retries
            }
        }while($true)
    }
    end
    {
        #restore progress preference
        $progressPreference = $pref
    }
}