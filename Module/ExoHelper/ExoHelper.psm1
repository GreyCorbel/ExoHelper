using namespace ExoHelper
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
New-AadAuthenticationFactory -ClientId (Get-ExoDefaultClientId) -TenantId 'mydomain.onmicrosoft.com' -AuthMode Interactive
New-ExoConnection -authenticationfactory $factory

Description
-----------
This command initializes connection to EXO REST API.
It uses instance of AADAuthenticationFactory for authentication with EXO REST API

.EXAMPLE
New-AadAuthenticationFactory -ClientId (Get-ExoDefaultClientId) -TenantId 'mydomain.onmicrosoft.com' -AuthMode Interactive | New-ExoConnection -IPPS

Description
-----------
This command initializes connection to IPPS REST API.
It uses instance of AADAuthenticationFactory for authentication with IPPS REST API passed via pipeline


#>
param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        #AAD authentication factory created via New-AadAuthenticationFactory
        #for user context, user factory created with clientId = fb78d390-0c51-40cd-8e17-fdbfab77341b (clientId of ExchangeOnlineManagement module) or your app with appropriate scopes assigned
        $AuthenticationFactory,
        
        [Parameter()]
        #Tenant ID when not the same as specified for factory - tenant native domain (xxx.onmicrosoft.com, or tenant GUID)
        [string]
        $TenantId,
        
        [Parameter()]
        #UPN of anchor mailbox
        #Default: UPN of caller or static system mailbox  (for app-only context)
        [string]
        $AnchorMailbox,

        [switch]
        #Connection is specialized to call IPPS commands
        #If not present, connection is specialized to call Exchange Online commands
        $IPPS
    )

    process
    {
        $Connection = [PSCustomObject]@{
            AuthenticationFactory = $AuthenticationFactory
            ConnectionId = [Guid]::NewGuid().ToString()
            TenantId = $null
            AnchorMailbox = $null
            ConnectionUri = $null
            IsIPPS = $IPPS.IsPresent
        }
        $claims = Get-ExoToken -Connection $Connection | Test-AadToken -PayloadOnly
        $Connection.TenantId = $claims.tid
        if($IPPS)
        {
            $Connection.ConnectionUri = "https://eur01b.ps.compliance.protection.outlook.com/AdminApi/beta/$($Connection.TenantId)/InvokeCommand"
        }
        else
        {
            $Connection.ConnectionUri = "https://outlook.office365.com/adminapi/beta/$($Connection.TenantId)/InvokeCommand"
        }

        if([string]::IsNullOrEmpty($AnchorMailbox))
        {
            if($null -ne $claims.upn)
            {
                #using caller's mailbox
                $Connection.AnchorMailbox = "UPN:$($claims.upn)"
            }
            else
            {
                #likely app-only context - use same static anchor mailbox as ExchangeOnlineManagement module uses
                $Connection.AnchorMailbox = "DiscoverySearchMailbox{D919BA05-46A6-415f-80AD-7E09334BB852}@$tenantId"
            }
        }
        else
        {
            $Connection.AnchorMailbox = "UPN:$AnchorMailbox"
        }

        $script:ConnectionContext = $Connection
        $script:ConnectionContext
    }
}

function Get-ExoDefaultClientId
{
    [CmdletBinding()]
    param ( )
    process
    {
        'fb78d390-0c51-40cd-8e17-fdbfab77341b'
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
        #Connection context as returned by New-ExoConnection
        #When not specified, uses most recently created connection context
        $Connection = $script:ConnectionContext,
        #Forces reauthentication
        #Usefiu when want to continue working after PIN JIT re-elevation
        [switch]$ForceRefresh
    )

    begin
    {
        if($null -eq $Connection )
        {
            throw 'Call New-ExoConnection first'
        }
    }
    process
    {
        if($Connection.IsIPPS)
        {
           $Scopes = "https://ps.compliance.protection.outlook.com/.default"
        }
        else
        {
            $Scopes = "https://outlook.office365.com/.default"
        }
        Get-AadToken -Factory $Connection.AuthenticationFactory -Scopes $scopes -ForceRefresh:$ForceRefresh
    }
}

function Encrypt-Value
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [AllowNull()]
        $UnsecureString
    )
    begin
    {
        $PublicKey = $MyInvocation.MyCommand.Module.PrivateData.Configuration.PublicKey
    }
    process
    {
        # Handling public key unavailability in client module for protection gracefully
        if ([string]::IsNullOrWhiteSpace($PublicKey))
        {
            # Error out if we are not in a position to protect the sensitive data before sending it over wire.
            throw 'Public key not found in the module definition';
        }

        if ($UnsecureString -ne '' -and $UnsecureString -ne $null)
        {
            $RSA = New-Object -TypeName System.Security.Cryptography.RSACryptoServiceProvider;
            $RSA.FromXmlString($PublicKey);
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($UnsecureString);
            $result = [byte[]]$RSA.Encrypt($bytes, $false);
            $RSA.Dispose();
            $result = [System.Convert]::ToBase64String($result);
            return $result;
        }
        return $UnsecureString;
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

.EXAMPLE
$connection = New-AadAuthenticationFactory -ClientId (Get-ExoDefaultClientId) -TenantId 'mydomain.onmicrosoft.com' -AuthMode Interactive | New-ExoConnection -IPPS
Invoke-ExoCommand -Connection $connection -Name 'Get-Label' -PropertiesToLoad 'ImmutableId','DisplayName' -RemoveOdataProperties -ShowWarnings

Description
-----------
This command creates connection for IPPS REST API, retrieves list of sensitivity labels returning only ImmutableId and DisplayName properties, and shows any warnings returned by IPPS REST API

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

        [Parameter()]
        [int]
            #Max results to return in single request
            #Default is 100
        $PageSize = 100,

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
        #Connection context as returned by New-ExoConnection
        #When not specified, uses most recently created connection context
        $Connection = $script:ConnectionContext

    )

    begin
    {
        $body = @{}
        if($PageSize -gt 1000)
        {
            $batchSize = 1000
        }
        if($PageSize -le 0)
        {
            $batchSize = 100
        }
        $batchSize = $pageSize

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
        $headers['X-AnchorMailbox'] = $Connection.anchorMailbox
        $headers['X-ClientApplication'] ='ExoHelper'

        #make sure that hashTable in parameters is properly decorated
        $keys = @()
        $Parameters.Keys | ForEach-Object { $keys += $_ }
        foreach($key in $Keys)
        {
            if($Parameters[$key] -is [hashtable])
            {
                $Parameters[$key]['@odata.type'] =  '#Exchange.GenericHashTable'
            }
            if($Parameters[$key] -is [System.Security.SecureString])
            {
                $cred = new-object System.Net.NetworkCredential -ArgumentList @($null, $Parameters[$key])
                $Parameters[$key] = Encrypt-Value -UnsecureString $cred.Password
            }
        }
        $body['CmdletInput'] = @{
            CmdletName = $Name
            Parameters = $Parameters
        }
        $retries = 0
        $resultsRetrieved = 0
        $pageUri = $uri
        $shouldContinue = $true #to support ErrorAction = SilentlyContinue that does not throw
        do
        {
            try {
                #new request id for each request
                $headers['client-request-id'] = [Guid]::NewGuid().ToString()
                #provide up to date token for each request of commands returning paged results that may take long to complete
                $headers['Authorization'] = (Get-ExoToken -Connection $Connection).CreateAuthorizationHeader()
                Write-Verbose "$([DateTime]::UtcNow.ToString('o'))`tResults:$resultsRetrieved`tRequestId: $($headers['client-request-id'])`tUri: $pageUri"
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
                        #disable progress bar that slows things down
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
                if(($PSVersionTable.psEdition -eq 'Desktop' -and $ex -is [System.Net.WebException]) -or ($PSVersionTable.psEdition -eq 'Core' -and $ex -is [Microsoft.PowerShell.Commands.HttpResponseException]))
                {
                    $responseHeaders = $ex.Response.Headers
                    $details = ($_.errordetails.message | ConvertFrom-Json).error
                    if($null -ne $details.details)
                    {
                        $errorData = $details.details.message.split('|')
                    }
                    else
                    {
                        $errorData = $details.message.split('|')
                    }
                    if($errorData.count -eq 3)
                    {
                        $ExoException = new-object ExoException -ArgumentList @($ex.Response.StatusCode, $errorData[0], $errorData[1], $errorData[2], $ex)
                    }
                    else
                    {
                        $ExoException = new-object ExoException -ArgumentList @($ex.Response.StatusCode, 'ExoGeneralError', $details.code, $details.message, $ex)
                    }

                    if($ex.response.statusCode -ne 429 -or $retries -ge $MaxRetries)
                    {
                        #different error or max retries exceeded
                        if($null -ne $exoException)
                        {
                            $shouldContinue = $false
                            throw $exoException
                        }
                        else
                        {
                            $shouldContinue = $false
                            throw
                        }
                    }
                }
                else
                {
                    #different exception type
                    $shouldContinue = $false
                    throw
                }
                $retries++
                if($ShowWarnings)
                {
                    Write-Warning "Retry #$retries"
                }
                else
                {
                    Write-Verbose "Retry #$retries"
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
        }while($null -ne $pageUri -and $resultsRetrieved -lt $ResultSize -and $shouldContinue)
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

Add-Type -TypeDefinition @'
    using System;
    using System.Net;
    namespace ExoHelper
    {
        public static class ExoHelperStringExtensions
        {
            public static long FromExoSize(this string input)
            {
                var start = input.IndexOf('(');
                var end = input.IndexOf(' ', start);
                long output = -1;
                long.TryParse(input.Substring(start + 1, end - start - 1).Replace(",", string.Empty), out output);
                return output;
            }
        }
        public class ExoException : Exception
        {
            public HttpStatusCode? StatusCode { get; set; }
            public string ExoErrorCode { get; set; }
            public string ExoErrorType { get; set; }
            public ExoException(HttpStatusCode? statusCode, string exoCode, string exoErrorType, string message):this(statusCode, exoCode, exoErrorType, message, null)
            {
            }
            public ExoException(HttpStatusCode? statusCode, string exoCode, string exoErrorType, string message, Exception innerException):base(message, innerException)
            {
                StatusCode = statusCode;
                ExoErrorCode = exoCode;
                ExoErrorType = exoErrorType;
            }
        }
    }
'@