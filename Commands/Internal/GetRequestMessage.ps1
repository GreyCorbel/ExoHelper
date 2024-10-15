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
