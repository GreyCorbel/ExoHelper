function Get-ExoException
{
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject]
        $ErrorRecord,
        [Parameter()]
        $httpCode
    )

    process
    {
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
