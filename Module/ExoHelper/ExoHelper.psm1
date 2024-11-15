#region Public commands
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
            #Default: DefaultRetryCount on EXO connection
        $MaxRetries = -1,

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

        [Parameter()]
        [System.Nullable[timespan]]
            #Timeout for the command execution
            #Default is timeout of the connection
            #If specified, must be lower than default connection timeout
            #See also https://makolyte.com/csharp-how-to-change-the-httpclient-timeout-per-request/ for more details on timeouts of http client
        $Timeout,

        [switch]
            #If we want to write any warnings returned by EXO REST API
        $ShowWarnings,

        [switch]
            #If we want to remove odata type descriptor properties from the output
        $RemoveOdataProperties,

        [switch]
            # If we want to include rate limits reported by REST API to verbose output
            # Requires verbose output to be enabled
        $ShowRateLimits,

        [Parameter()]
        #Connection context as returned by New-ExoConnection
        #When not specified, uses most recently created connection context
        $Connection = $script:ConnectionContext
    )

    begin
    {
        if($null -ne $Timeout)
        {
            $cts = new-object System.Threading.CancellationTokenSource($Timeout)
        }
        else {
            $cts = new-object System.Threading.CancellationTokenSource($Connection.HttpClient.Timeout)
        }

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
        if($MaxRetries -eq -1)
        {
            $MaxRetries = $Connection.DefaultRetryCount
        }
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
                $Parameters[$key] = EncryptValue -UnsecureString $cred.Password
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
                $requestMessage = GetRequestMessage -Uri $pageUri -Headers $headers -Body ($body | ConvertTo-Json -Depth 9)
                $response = $Connection.HttpClient.SendAsync($requestMessage, $cts.Token).GetAwaiter().GetResult()
                $requestMessage.Dispose()

                if($null -ne $response.Content -and $response.StatusCode -ne [System.Net.HttpStatusCode]::NoContent)
                {
                    $payload = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                    if($response.content.Headers.ContentType.MediaType -eq 'application/json')
                    {
                        try {
                            $responseData = $Payload | ConvertFrom-Json
                        }
                        catch {
                            Write-Warning "Received unexpected non-JSON response: $payload"
                            $responseData = $payload
                        }
                    }
                    else
                    {
                        Write-Warning "Received non-JSON response: $($response.content.Headers.ContentType.MediaType)"
                        $responseData = $payload
                    }
                }
                else
                {
                    $responseData = $null
                }
                if($response.IsSuccessStatusCode)
                {
                    if($null -ne $responseData)
                    {
                        if($responseData -is [string])
                        {
                            #we did not receive JSON response - return it and finish
                            $shouldContinue = $false
                            $responseData
                        }
                        else
                        {
                            #we have parsed json response
                            if($ShowWarnings -and $null -ne $responseData.'@adminapi.warnings')
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
                            Write-Verbose "$([DateTime]::UtcNow.ToString('o'))`tResults:$resultsRetrieved`tRequestId: $($headers['client-request-id'])`tUri: $pageUri"
                            $pageUri = $responseData.'@odata.nextLink'
                        }
                    }
                }
                else
                {
                    #request failed
                    $ex = $null
                    if($response.StatusCode -ne [System.Net.HttpStatusCode]::TooManyRequests)
                    {
                        $shouldContinue = $false
                        $ex = $responseData | Get-ExoException -httpCode $response.StatusCode
                    }
                    else
                    {
                        if($retries -ge $MaxRetries)
                        {
                            $shouldContinue = $false
                            $ex = New-Object ExoHelper.ExoException($response.StatusCode, 'ExoTooManyRequests', '', 'Max retry count for throttled request exceeded')
                        }
                    }
                    if($null -ne $ex)
                    {
                        Write-Verbose "Handling exception: StatusCode $($ex.StatusCode)`tExoErrorCode: $($ex.ExoErrorCode)`tExoErrorType: $($ex.ExoErrorType)"
                        switch($ErrorActionPreference)
                        {
                            'Stop' { throw $ex }
                            'Continue' { Write-Error -Exception $ex; return }
                            default { return }
                        }
                    }
                    #TooManyRequests --> let's wait and retry
                    $retries++
                    switch($WarningPreference)
                    {
                        'Continue' { Write-Warning "Retry #$retries" }
                        'SilentlyContinue' { Write-Verbose "Retry #$retries" }
                    }

                    #wait some time
                    Start-Sleep -Seconds $retries
                }
            }
            catch
            {
                $shouldContinue = $false
                throw
            }
            finally
            {
                if($ShowRateLimits)
                {
                    $val = $null
                    if($response.Headers.TryGetValues('Rate-Limit-Remaining', [ref]$val)) 
                    {
                        $rateLimitRemaining = $val
                        if($response.Headers.TryGetValues('Rate-Limit-Reset', [ref]$val))
                        {
                            $rateLimitReset = $val
                            Write-Verbose "Rate limit remaining: $rateLimitRemaining`tRate limit reset: $rateLimitReset"
                        }
                    }
                }
            }
        }while($null -ne $pageUri -and ($resultsRetrieved -lt $ResultSize) -and $shouldContinue)
    }
    end
    {
        if($null -ne $cts)
        {
            $cts.Dispose()
        }
    }
}
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
$factory = New-AadAuthenticationFactory -ClientId (Get-ExoDefaultClientId) -TenantId 'mydomain.onmicrosoft.com' -AuthMode WAM
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
        #Tenant ID when not the same as specified for factory
        #Must be tenant native domain (xxx.onmicrosoft.com)
        [string]
        $TenantId,
        
        [Parameter()]
        #UPN of anchor mailbox
        #Default: UPN of caller or static system mailbox  (for app-only context)
        [string]
        $AnchorMailbox,

        [Parameter()]
        [timespan]
            #Default timeout for the EXO command execution
        $DefaultTimeout = [timespan]::FromMinutes(60),

        [Parameter()]
        [int]
            #Default retry count for the EXO command execution
        $DefaultRetryCount = 10,

        [switch]
        #Connection is specialized to call IPPS commands
        #If not present, connection is specialized to call Exchange Online commands
        $IPPS
    )

    process
    {
        $Connection = [PSCustomObject]@{
            PSTypeName = "ExoHelper.Connection"
            AuthenticationFactory = $AuthenticationFactory
            ConnectionId = [Guid]::NewGuid().ToString()
            TenantId = $null
            AnchorMailbox = $null
            ConnectionUri = $null
            IsIPPS = $IPPS.IsPresent
            HttpClient = new-object System.Net.Http.HttpClient
            DefaultRetryCount = $DefaultRetryCount
        }
        $Connection.HttpClient.DefaultRequestHeaders.Add("User-Agent", "ExoHelper")
        $Connection.HttpClient.Timeout = $DefaultTimeout

        #explicitly authenticate when establishing connection to catch any authentication problems early
        Get-ExoToken -Connection $Connection | Out-Null
        if([string]::IsNullOrEmpty($TenantId))
        {
            $TenantId = $AuthenticationFactory.TenantId
        }
        if([string]::IsNullOrEmpty($TenantId))
        {
            throw (new-object ExoHelper.ExoException([System.Net.HttpStatusCode]::BadRequest, 'ExoMissingTenantId', 'ExoInitializationError', 'TenantId is not specified and cannot be determined automatically - please specify TenantId parameter'))
        }
        $Connection.TenantId = $TenantId
        
        if($IPPS)
        {
            $Connection.ConnectionUri = "https://eur02b.ps.compliance.protection.outlook.com/adminapi/beta/$($Connection.TenantId)/InvokeCommand"
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
                $Connection.AnchorMailbox = "UPN:DiscoverySearchMailbox{D919BA05-46A6-415f-80AD-7E09334BB852}@$tenantId"
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
#endregion Public commands
#region Internal commands
#encrypts data using MS provided public key
#key stored in module private data
#MS rotates the key regularly; at least one version back is supported
function EncryptValue
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowEmptyString()]
        [AllowNull()]
        $UnsecureString,
        [Parameter()]
        [string]$Key = $script:PublicKey
    )
    process
    {
        # Handling public key unavailability in client module for protection gracefully
        if ([string]::IsNullOrWhiteSpace($Key))
        {
            # Error out if we are not in a position to protect the sensitive data before sending it over wire.
            throw 'Public key not loaded. Cannot encrypt sensitive data.';
        }

        if (-not [string]::IsNullOrWhiteSpace($UnsecureString))
        {
            $RSA = New-Object -TypeName System.Security.Cryptography.RSACryptoServiceProvider;
            $RSA.FromXmlString($Key);
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($UnsecureString);
            $result = [byte[]]$RSA.Encrypt($bytes, $false);
            $RSA.Dispose();
            $result = [System.Convert]::ToBase64String($result);
            return $result;
        }
        return $UnsecureString;
    }
}
function Get-ExoException
{
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]
        $ErrorRecord,
        [Parameter()]
        $httpCode
    )

    process
    {
        if($ErrorRecord -is [string])
        {
            return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoErrorWithPlainText', '', $ErrorRecord)
        }
        #structured error
        if($null -ne $errorRecord.error.details.message)
        {
            $message = $errorRecord.error.details.message
            $errorData = $message.split('|')
            if($errorData.count -eq 3)
            {
                return new-object ExoHelper.ExoException -ArgumentList @($httpCode, $errorData[0], $errorData[1], $errorData[2])
            }
            else
            {
                return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoErrorWithUnknownDetail', '', $message)
            }
        }
        if($null -ne $errorRecord.error.innerError.internalException)
        {
            return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoErrorWithInternalException', $errorRecord.error.innerError.type, $errorRecord.error.innerError.internalException.Message)
        }
        if($null -ne $errorRecord.error)
        {
            return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoErrorWithMissingDetail', '', "$($errorRecord.error)")
        }
        return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoUnknownError', '', "$($errorRecord.error)")
    }
}
function GetRequestMessage
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [string]
        $Uri,
        [Parameter()]
        [hashtable]
        $Headers,
        [Parameter()]
        [string]
        $Body
    )
    begin
    {
        $ContentType = 'application/json'
        $Method = [System.Net.Http.HttpMethod]::Post
    }
    process
    {
        $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::$Method, (new-object System.Uri($Uri)))
        if($null -ne $Headers)
        {
            foreach($header in $Headers.Keys)
            {
                $request.Headers.TryAddWithoutValidation($header, $Headers[$header]) | Out-Null
            }
        }
        if($null -ne $Body)
        {
            $payload = ([System.Text.Encoding]::UTF8.GetBytes($Body));
            $request.Content = [System.Net.Http.ByteArrayContent]::new($payload)
            $request.Content.headers.Add('Content-Encoding','utf-8')
            $request.Content.headers.Add('Content-Type','application/json')
        }
        $request.Method = $Method
        $request
    }
}
Function Init
{
    param()

    process
    {
        #Add JIT compiled helpers. Load only if not loaded previously
        $referencedAssemblies=@()
        $helpers = 'ExoException', 'StringExtensions'
        foreach($helper in $helpers)
        {
            #compiled helpers are in ExoHelper namespace
            if($null -eq ("ExoHelper.$helper" -as [type]))
            {
                $helperDefinition = Get-Content "$PSScriptRoot\Helpers\$helper.cs" -Raw
                Add-Type -TypeDefinition $helperDefinition -ReferencedAssemblies $referencedAssemblies -WarningAction SilentlyContinue -IgnoreWarnings
            }
        }
        
        #refresh cached public key, if needed
        $PublicKeyConfig = $MyInvocation.MyCommand.Module.PrivateData.Configuration.ExoPublicKey
        $cacheFile = [System.IO.Path]::Combine($env:TEMP, $PublicKeyConfig.LocalFile)
        $needsRefresh = $false

        if(-not [System.IO.File]::Exists($cacheFile))
        {
            $needsRefresh = $true
        }
        else {
            # local file exists
            $fileInfo = [System.IO.FileInfo]::new($cacheFile)
            if($fileInfo.LastWriteTime -lt (Get-Date).AddDays(-7))
            {
                $needsRefresh = $true
            }
        }
        if($needsRefresh)
        {
            try
            {
                Invoke-WebRequest -Uri $PublicKeyConfig.Link -OutFile $cacheFile -ErrorAction Stop
            }
            catch
            {
                write-warning 'Local copy of public key file is ooutdated or does not exist and failed to download public key. Module may not work correctly.'
                $script:PublicKey = $null
                return
            }
        }
        $script:PublicKey = [System.IO.File]::ReadAllText($cacheFile)
    }
}
#removes odata type descriptor properties from the object
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
#endregion Internal commands
#region Module initialization
Init
#endregion Module initialization
