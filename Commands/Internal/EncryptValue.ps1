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
