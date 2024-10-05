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
