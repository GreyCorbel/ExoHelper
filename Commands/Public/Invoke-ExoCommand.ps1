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

        [Parameter()]
            #Status codes that are considered retryable
        [System.Net.HttpStatusCode[]]$RetryableStatusCodes,
            
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
        if($RetryableStatusCodes.Count -eq 0)
        {
            $RetryableStatusCodes = @('ServiceUnavailable', 'GatewayTimeout', 'RequestTimeout')
            if($PSVersionTable.PSEdition -eq 'Core')
            {
                #this is not available in .NET Frameworl
                $RetryableStatusCodes += 'TooManyRequests'
            }
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
                            Write-Warning "Received unexpected non-JSON response with http staus $($response.StatusCode): $payload"
                            $responseData = $payload
                        }
                    }
                    else
                    {
                        Write-Warning "Received non-JSON response with http status $($response.StatusCode): $payload"
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
                    $exceptionType = $null
                    if($null -ne $response.Headers)
                    {
                        $response.Headers.TryGetValues('X-ExceptionType', [ref]$exceptionType) | out-null
                    }
                    <# if($response.StatusCode -notin $RetryableStatusCodes `
                        -and $exceptionType -notin @('UnableToWriteToAadException') `
                        -and $responseData -notlike 'You have reached the maximum number of concurrent requests per tenant. Please wait and try again*' `
                        -and $responseData -notlike '*issue may be transient*' `
                        ) #>
                    if($response.StatusCode -notin $RetryableStatusCodes -and $responseData -notlike 'You have reached the maximum number of concurrent requests per tenant. Please wait and try again*' )
                    {
                        $shouldContinue = $false
                        if($null -ne $responseData)
                        {
                            #we have structured error
                            $ex = $responseData | Get-ExoException -httpCode $response.StatusCode -exceptionType $exceptionType
                        }
                        else
                        {
                            #we have plain text error
                            $ex = new-object ExoHelper.ExoException($response.StatusCode, 'ExoErrorWithPlainText', $exceptionType, $payload)
                        }
                    }
                    else
                    {
                        #we wait on http 429 or throttling message
                        if($retries -ge $MaxRetries)
                        {
                            $shouldContinue = $false
                            if([string]::IsNullOrEmpty($payload))
                            {
                                $ex = New-Object ExoHelper.ExoException($response.StatusCode, 'ExoTooManyRequests', $exceptionType, 'Max retry count for request exceeded')
                            }
                            else
                            {
                                $ex = New-Object ExoHelper.ExoException($response.StatusCode, 'ExoTooManyRequests', $exceptionType, $payload)
                            }
                        }
                    }
                    if($null -ne $ex)
                    {
                        switch($ErrorActionPreference)
                        {
                            'Stop' { throw $ex }
                            'Continue' { Write-Error -Exception $ex; return }
                            default { return }
                        }
                    }
                    #TooManyRequests --> let's wait and retry
                    #$headers = @{}
                    #$response.Headers | ForEach-Object { $headers[$_.Key] = $_.Value }
                    #$headersObject = [PSCustomObject]$headers

                    $retries++
                    $val = $null
                    if($null -ne $response.Headers -and $response.Headers.TryGetValues('Retry-After', [ref]$val))
                    {
                        $retryAfter = [int]($val[0])
                    }
                    else
                    {
                        $retryAfter = 3 * $retries
                    }

                    switch($WarningPreference)
                    {
                        'Continue' { 
                            Write-Warning "Retry #$retries after $retryAfter seconds"
                            break;
                        }
                        'SilentlyContinue' {
                            Write-Verbose "Retry #$retries after $retryAfter seconds"
                            break;
                        }
                    }

                    #wait some time
                    Start-Sleep -Seconds $retryAfter
                }
            }
            catch
            {
                $shouldContinue = $false
                throw
            }
            finally
            {
                if($null -ne $response)
                {
                    if($ShowRateLimits -and $null -ne $response.Headers)
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
                    $response.Dispose()
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
