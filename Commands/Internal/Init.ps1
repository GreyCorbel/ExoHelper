Function Init
{
    param()

    process
    {
        #Add JIT compiled helpers. Load only if not loaded previously
        $referencedAssemblies=@()
        $helpers = 'ExoException', 'StringExtensions'
        foreach($helper in $helpers)
        {
            #compiled helpers are in ExoHelper namespace
            if($null -eq ("ExoHelper.$helper" -as [type]))
            {
                $helperDefinition = Get-Content "$PSScriptRoot\Helpers\$helper.cs" -Raw
                Add-Type -TypeDefinition $helperDefinition -ReferencedAssemblies $referencedAssemblies -WarningAction SilentlyContinue -IgnoreWarnings
            }
        }
        
        #refresh cached public key, if needed
        $PublicKeyConfig = $MyInvocation.MyCommand.Module.PrivateData.Configuration.ExoPublicKey
        $cacheFile = [System.IO.Path]::Combine($env:TEMP, $PublicKeyConfig.LocalFile)
        $needsRefresh = $false

        if(-not [System.IO.File]::Exists($cacheFile))
        {
            $needsRefresh = $true
        }
        else {
            # local file exists
            $fileInfo = [System.IO.FileInfo]::new($cacheFile)
            if($fileInfo.LastWriteTime -lt (Get-Date).AddDays(-7))
            {
                $needsRefresh = $true
            }
        }
        if($needsRefresh)
        {
            try
            {
                Invoke-WebRequest -Uri $PublicKeyConfig.Link -OutFile $cacheFile -ErrorAction Stop
            }
            catch
            {
                write-warning 'Local copy of public key file is ooutdated or does not exist and failed to download public key. Module may not work correctly.'
                $script:PublicKey = $null
                return
            }
        }
        $script:PublicKey = [System.IO.File]::ReadAllText($cacheFile)
    }
}
