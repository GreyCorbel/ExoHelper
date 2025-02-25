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
        if($authenticationFactory -is [string])
        {
            $f = Get-AadAuthenticationFactory -Name $authenticationFactory
            if($null -eq $f)
            {
                throw (new-object ExoHelper.ExoException([System.Net.HttpStatusCode]::BadRequest, 'ExoMissingAuthenticationFactory', 'ExoInitializationError', "Factory with name $authenticationFactory not found"))
            }
        }
        else
        {
            $f = $authenticationFactory
        }
        if([string]::IsNullOrEmpty($tenantId) )
        {
            $tenantId = $f.tenantId
        }
        if([string]::IsNullOrEmpty($TenantId))
        {
            throw (new-object ExoHelper.ExoException([System.Net.HttpStatusCode]::BadRequest, 'ExoMissingTenantId', 'ExoInitializationError', 'TenantId is not specified and cannot be determined automatically - please specify TenantId parameter'))
        }

        $Connection = [PSCustomObject]@{
            PSTypeName = "ExoHelper.Connection"
            AuthenticationFactory = $f
            ConnectionId = [Guid]::NewGuid().ToString()
            TenantId = $tenantId
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
