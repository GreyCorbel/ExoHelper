function New-ExoConnection
{
<#
.SYNOPSIS
    Initializes EXO connection

.DESCRIPTION
    Initializes EXO connection

.OUTPUTS
    None

.EXAMPLE
New-ExoConnection -authenticationfactory $factory -TenantId mydomain.onmicrosoft.com

Description
-----------
This command initializes connection to EXO REST API.
It uses instance of AADAuthenticationFactory for authentication with EXO REST API
#>
param
    (
        [Parameter(Mandatory)]
        #AAD authentication factory created via New-AadAuthenticationFactory
        #for user context, user factory created with clientId = fb78d390-0c51-40cd-8e17-fdbfab77341b (clientId of ExchangeOnlineManagement module) or your app with appropriate scopes assigned
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
        $Connection = [PSCustomObject]@{
            AuthenticationFactory = $AuthenticationFactory
            TenantId = $tenantId
            AnchorMailbox = "UPN:$anchorMailbox"
            ConnectionId = [Guid]::NewGuid().ToString()
            ConnectionUri = "https://outlook.office365.com/adminapi/beta/$tenantId/InvokeCommand"
        }

        if([string]::IsNullOrEmpty($AnchorMailbox))
        {
            $claims = Get-ExoToken -Connection $Connection | Test-AadToken -PayloadOnly
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
        $Connection.AnchorMailbox = "UPN:$anchorMailbox"
        $script:ConnectionContext = $Connection
        $script:ConnectionContext
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
        [Parameter()]
        $Connection = $script:ConnectionContext
    )

    process
    {
        Get-AadToken -Factory $Connection.AuthenticationFactory -Scopes 'https://outlook.office365.com/.default'
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

        [Parameter()]
        [int]
            #Max results to return
            #1000 is a minimum, and min increment is 1000
        $ResultSize = [int]::MaxValue,

        [switch]
            #If we want to write any warnings returned by EXO REST API
        $ShowWarnings,

        [switch]
            #If we want to remove odata type descriptor properties from the output
        $RemoveOdataProperties,

        [switch]
        #If we want to include rate limits reported by REST API to verbose output
        $ShowRateLimits,

        [Parameter()]
        $Connection = $script:ConnectionContext

    )

    begin
    {
        $body = @{}
        $batchSize = 100
        $uri = $Connection.ConnectionUri
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
        $headers = @{}
        $headers['X-CmdletName'] = $Name
        $headers['Prefer'] = "odata.maxpagesize=$batchSize"
        $headers['connection-id'] = $Connection.connectionId
        $headers['X-AnchorMailbox'] =$Connection.anchorMailbox

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
        $resultsRetrieved = 0
        $pageUri = $uri
        do
        {
            do
            {
                try {
                    #new request id for each request
                    $headers['client-request-id'] = [Guid]::NewGuid().ToString()
                    #provide up to date token for each request of commands returning paged results that may take long to complete
                    $headers['Authorization'] = (Get-ExoToken -Connection $Connection).CreateAuthorizationHeader()
                    Write-Verbose "RequestId: $($headers['client-request-id'])`tUri: $pageUri"
                    $splat = @{
                        Uri = $pageUri
                        Method = 'Post'
                        Body = ($body | ConvertTo-Json -Depth 9)
                        Headers = $headers
                        ContentType = 'application/json'
                        ErrorAction = 'Stop'
                        Verbose = $false
                    }
                    #add edition-specific parameters
                    if($PSEdition -eq 'Desktop')
                    {
                        $splat['UseBasicParsing'] = $true
                    }
                    else
                    {
                        if($psversionTable.PSVersion -gt '7.4') #7.4+ supports ProgressAction
                        {
                            $splat['ProgressAction'] = 'SilentlyContinue'
                        }
                    }
                    $response = Invoke-WebRequest @splat
                    #we may process the headers in the future to see rate limit remaining, etc.
                    $responseHeaders = $response.Headers
    
                    $responseData = $response.Content | ConvertFrom-Json
                    
                    if($ShowWarnings)
                    {
                        foreach($warning in $responseData.'@adminapi.warnings')
                        {
                            Write-Warning $warning
                        }
                    }
                    $resultsRetrieved+=$responseData.value.Count
                    if($RemoveOdataProperties)
                    {
                        $responseData.value | RemoveExoOdataProperties
                    }
                    else {
                        $responseData.value
                    }
                    $pageUri = $responseData.'@odata.nextLink'
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
                    $responseHeaders = $ex.Response.Headers
                    $retries++
                    if($ShowWarnings)
                    {
                        Write-Warning "Retry #$retries"
                    }
                    else
                    {
                        Write-Verbose "Retry #$retries"
                    }
                    if($retries -gt $MaxRetries)
                    {
                        #max retries exhausted
                        throw
                    }
                    #wait some time
                    Start-Sleep -Seconds $retries
                }
                finally
                {
                    if($ShowRateLimits)
                    {
                        if($null -ne $responseHeaders -and $null -ne $responseHeaders['Rate-Limit-Remaining'] -and $null -ne $responseHeaders['Rate-Limit-Reset'])
                        {
                            if($PSVersionTable.psEdition -eq 'Desktop')
                            {
                                Write-Verbose "Rate limit remaining: $($responseHeaders['Rate-Limit-Remaining'])`tRate limit reset: $($responseHeaders['Rate-Limit-Reset'])"
                            }
                            else
                            {
                                #Core
                                Write-Verbose "Rate limit remaining: $($responseHeaders['Rate-Limit-Remaining'][0])`tRate limit reset: $($responseHeaders['Rate-Limit-Reset'][0])"
                            }
                        }
                    }
                }
            }while($null -ne $pageUri -and $resultsRetrieved -le $ResultSize)
            break
        }while($true)
    }
    end
    {
        #restore progress preference
        $progressPreference = $pref
    }
}

function RemoveExoOdataProperties
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject]
        $Object
    )
    begin
    {
        $propsToRemove = $null
    }
    process
    {
        if($null -eq $propsToRemove)
        {
            $propsToRemove = $Object.PSObject.Properties | Where-Object { $_.Name.IndexOf('@') -ge 0 }
        }
        foreach($prop in $propsToRemove)
        {
            $Object.PSObject.Properties.Remove($prop.Name)
        }
        $Object
    }
}