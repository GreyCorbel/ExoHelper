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
            $request.Content = [System.Net.Http.StringContent]::new($Body, [System.Text.Encoding]::UTF8, $ContentType)
        }
        $request.Method = $Method
        $request
    }
}
