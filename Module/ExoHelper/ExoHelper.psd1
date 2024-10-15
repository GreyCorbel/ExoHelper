#
# Module manifest for module 'ExoHelper'
#
# Generated by: JiriFormacek
#
# Generated on: 1/3/2024
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'ExoHelper.psm1'

# Version number of this module.
ModuleVersion = '3.0.0'

# Supported PSEditions
CompatiblePSEditions = @('Desktop', 'Core')

# ID used to uniquely identify this module
GUID = 'c113d53c-8c32-4ebf-998f-16ef76b04cd9'

# Author of this module
Author = 'Jiri Formacek'

# Company or vendor of this module
CompanyName = 'GreyCorbel Solutions'

# Copyright statement for this module
Copyright = '(c) Jiri Formacek. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Simple wrapper module that directly calls EXO REST API without the need for full heavy-weight ExchangeOnlineManagement module'

# Minimum version of the PowerShell engine required by this module
PowerShellVersion = '5.1'

# Name of the PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# ClrVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @(@{ModuleName="AadAuthenticationFactory"; ModuleVersion="3.1.1"; GUID='9d860f96-4bde-41d3-890b-1a3f51c34d68'})

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
FormatsToProcess = @('ExoHelper.format.ps1xml')

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @('New-ExoConnection', 'Get-ExoDefaultClientId', 'Invoke-ExoCommand', 'Get-ExoToken')

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @('ExchangeOnlineManagement','PSEdition_Core','PSEdition_Desktop')

        # A URL to the license for this module.
        LicenseUri = 'https://raw.githubusercontent.com/GreyCorbel/ExoHelper/main/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/greycorbel/ExoHelper'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

        # Prerelease string of this module
        Prerelease = 'beta2'

        # Flag to indicate whether the module requires explicit user acceptance for install/update/save
        # RequireLicenseAcceptance = $false

        # External dependent modules of this module
        # ExternalModuleDependencies = @()

    } # End of PSData hashtable
    Configuration = @{
        # Microsoft-generated public key that protects sensitive information sent to Exo REST API
        # Rotated regularly and few version back are supported
        # Takem from downloaded temp module that comes as a result of Connect-ExchangeOnline
        ExoPublicKey = @{
            Link = 'https://raw.githubusercontent.com/GreyCorbel/ExoHelper/main/PublicKey.xml'
            LocalFile = "ExoHelper_PublicKey.xml"
        }
    }
} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

