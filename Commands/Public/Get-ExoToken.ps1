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
