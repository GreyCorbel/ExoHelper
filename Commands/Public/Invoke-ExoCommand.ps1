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
                Write-Verbose "$([DateTime]::UtcNow.ToString('o'))`tResults:$resultsRetrieved`tRequestId: $($headers['client-request-id'])`tUri: $pageUri"
                $requestMessage = GetRequestMessage -Uri $pageUri -Headers $headers -Body ($body | ConvertTo-Json -Depth 9)
                $response = $Connection.HttpClient.SendAsync($requestMessage).GetAwaiter().GetResult()
                $requestMessage.Dispose()

                if($null -ne $response.Content -and $response.StatusCode -ne [System.Net.HttpStatusCode]::NoContent)
                {
                    $payload = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $responseData = $Payload | ConvertFrom-Json
                }
                else
                {
                    $responseData = $null
                }
                if($response.IsSuccessStatusCode)
                {
                    if($null -ne $responseData)
                    {

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
                        $pageUri = $responseData.'@odata.nextLink'
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
}
