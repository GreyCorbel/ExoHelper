function Get-ExoException
{
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]
        $ErrorRecord,
        [Parameter()]
        $httpCode,
        [Parameter()]
        $exceptionType
    )

    process
    {
        if($ErrorRecord -is [string])
        {
            if([string]::IsNullOrEmpty($exceptionType))
            {
                return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoErrorWithPlainText', '', $ErrorRecord)
            }
            else {
                return new-object ExoHelper.ExoException -ArgumentList @($httpCode, $exceptionType, '', $ErrorRecord)
            }
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
                return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoErrorWithUnknownDetail', $exceptionType, $message)
            }
        }
        if($null -ne $errorRecord.error.innerError.internalException)
        {
            return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoErrorWithInternalException', $errorRecord.error.innerError.type, $errorRecord.error.innerError.internalException.Message)
        }
        if($null -ne $errorRecord.error)
        {
            return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoErrorWithMissingDetail', $exceptionType, "$($errorRecord.error)")
        }
        return new-object ExoHelper.ExoException -ArgumentList @($httpCode, 'ExoUnknownError', $exceptionType, "$($errorRecord.error)")
    }
}
