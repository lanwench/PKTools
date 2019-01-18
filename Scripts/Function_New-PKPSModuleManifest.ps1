﻿#requires -Version 3
Function New-PKPSModuleManifest {
<#
.SYNOPSIS 
    Creates a new PowerShell module manifest using New-ModuleManifest

.DESCRIPTION
    Creates a new PowerShell module manifest
    Supports ShouldProcess
    Outputs a PSObject

.OUTPUTS
    PSObject

.NOTES        
    Name    : Function_New-PKPSModuleManifest.ps1 
    Author  : Paula Kingsley
    Created : 2018-12-17
    Version : 01.00.1000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
    
        v01.00.0000 - 2018-12017 - Created script

.EXAMPLE
    PS C:\> New-PKPSModuleManifest -ModuleName MyNewModule -ModulePath C:\Users\jbloggs\git\Modules\MyNewModule -Author "Joe Bloggs" -CompanyName "MegaCorp International, Inc." -Description "My new test module" -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                           
        ---                   -----                                           
        ModuleName            MyNewModule                               
        ModulePath            C:\Users\jbloggs\git\Modules\MyNewModule
        Author                Joe Bloggs                                  
        CompanyName           MegaCorp International, Inc.                     
        Description           My new test module
        Verbose               True                                            
        CopyrightDate         2018                                            
        MininumPSVersion      3.0.0                                           
        ModuleVersion         1.0.0                                           
        SuppressConsoleOutput False                                           
        ScriptName            New-PKPSModuleManifest                          
        ScriptVersion         1.0.0                                           

        Action: Create new PowerShell module manifest
        Create new module manifest with the following specifications:

        Name              Value                                                                  
        ----              -----                                                                  
        Copyright         Copyright 2018, MegaCorp International, Inc.                            
        Path              C:\Users\jbloggs\git\Modules\MyNewModule\MyNewModule.psd1
        Verbose           True                                                                   
        CompanyName       MegaCorp International, Inc.                                            
        Description       My new test module
        Author            Joe Bloggs                                                         
        Passthru          True                                                                   
        RootModule        MyNewModule                                                      
        ErrorAction       Stop                                                                   
        PowershellVersion                                                                        



        VERBOSE: Performing the operation "Creating the "C:\Users\jbloggs\git\Modules\MyNewModule\MyNewModule.psd1" module manifest file." on target "C:\Users\jbloggs\git\Modules\MyNewModule\MyNewModule.psd1".
        #
        # Module manifest for module 'MyNewModule'
        #
        # Generated by: Joe Bloggs
        #
        # Generated on: 2018-12-17
        #

        @{

        # Script module or binary module file associated with this manifest.
        RootModule = 'MyNewModule'

        # Version number of this module.
        ModuleVersion = '1.0'

        # Supported PSEditions
        # CompatiblePSEditions = @()

        # ID used to uniquely identify this module
        GUID = '6a0cbc93-ebe7-454a-b717-a0a3963d52df'

        # Author of this module
        Author = 'Joe Bloggs'

        # Company or vendor of this module
        CompanyName = 'MegaCorp International, Inc.'

        # Copyright statement for this module
        Copyright = 'Copyright 2018, MegaCorp International, Inc.'

        # Description of the functionality provided by this module
        Description = 'My new test module'

        # Minimum version of the Windows PowerShell engine required by this module
        # PowerShellVersion = ''

        # Name of the Windows PowerShell host required by this module
        # PowerShellHostName = ''

        # Minimum version of the Windows PowerShell host required by this module
        # PowerShellHostVersion = ''

        # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
        # DotNetFrameworkVersion = ''

        # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
        # CLRVersion = ''

        # Processor architecture (None, X86, Amd64) required by this module
        # ProcessorArchitecture = ''

        # Modules that must be imported into the global environment prior to importing this module
        # RequiredModules = @()

        # Assemblies that must be loaded prior to importing this module
        # RequiredAssemblies = @()

        # Script files (.ps1) that are run in the caller's environment prior to importing this module.
        # ScriptsToProcess = @()

        # Type files (.ps1xml) to be loaded when importing this module
        # TypesToProcess = @()

        # Format files (.ps1xml) to be loaded when importing this module
        # FormatsToProcess = @()

        # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
        # NestedModules = @()

        # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
        FunctionsToExport = '*'

        # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
        CmdletsToExport = '*'

        # Variables to export from this module
        VariablesToExport = '*'

        # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
        AliasesToExport = '*'

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
                # Tags = @()

                # A URL to the license for this module.
                # LicenseUri = ''

                # A URL to the main website for this project.
                # ProjectUri = ''

                # A URL to an icon representing this module.
                # IconUri = ''

                # ReleaseNotes of this module
                # ReleaseNotes = ''

            } # End of PSData hashtable

        } # End of PrivateData hashtable

        # HelpInfo URI of this module
        # HelpInfoURI = ''

        # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
        # DefaultCommandPrefix = ''

        }

        VERBOSE: Successfully created new PowerShell module manifest for 'MyNewModule' in 'C:\Users\jbloggs\git\Modules\MyNewModule'


#>
[CmdletBinding(
    SupportsShouldProcess=$True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Mandatory = $True,
        HelpMessage = "Module name (e.g., MyNewPSModule)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$ModuleName,

    [Parameter(
        Mandatory = $True,
        HelpMessage = "Absolute path to module (e.g., c:\users\jbloggs\repo\modules\MyNewPSModule)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        If ($Null = Test-Path $_) {$True}
        Else {$False}
    })]
    [string]$ModulePath,

    [Parameter(
        HelpMessage = "Author's name (default is `$Env:CurrentUser)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Author = $env:USERNAME,

    [Parameter(
        HelpMessage = "Company name (default is none)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyName,

    [Parameter(
        Mandatory = $True,
        HelpMessage = "Description (e.g., 'Tools for making things'"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Description,

    [Parameter(
        HelpMessage = "Copyright date (default is yyyy)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$CopyrightDate = $(get-date -f yyyy),

    [Parameter(
        HelpMessage = "Minimum PowerShell version supported (default is 3.0.0)"
    )]
    [ValidateScript({
        If ($_ -as [version]) {$True}
        Else {$False}
    })]
    [ValidateNotNullOrEmpty()]
    [string]$MininumPSVersion = "3.0.0",

    [Parameter(
        HelpMessage = "Module version (default is 1.0.0)"
    )]
    [ValidateScript({
        If ($_ -as [version]) {$True}
        Else {$False}
    })]
    [ValidateNotNullOrEmpty()]
    [string]$ModuleVersion = "1.0.0",

    [Parameter(
        HelpMessage = "Suppress all non-verbose console output"
    )]
    [switch] $SuppressConsoleOutput

)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where{Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"

    # Just in case
    $Msg = "This function uses New-ModuleManifest; you might consider looking at Plaster instead!`n`t* https://github.com/PowerShell/Plaster`n`t* https://kevinmarquette.github.io/2017-05-12-Powershell-Plaster-adventures-in/`n`n"
    Write-Verbose $Msg

    # For console
    $Activity = "Create new PowerShell module manifest"
    $Msg = "Action: $Activity"
    $BGColor = $Host.UI.RawUI.BackgroundColor
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}
Process {

    # Make sure this doesn't already exist
    If ($Null = Test-Path "$ModulePath\$ModuleName.psd1") {
        $Msg = "$ModulePath\$ModuleName.psd1 already exists"
        $Host.UI.WriteErrorLine("ERROR: $Msg")
    }
    Else {
        # Splat for New-ModuleManifest
        $Param_NewModule = @{}
        $Param_NewModule = @{
            Path              = "$ModulePath\$ModuleName.psd1"
            Author            = $Author
            Copyright         = $Null
            Description       = $Description
            PowershellVersion = $PSVersion
            RootModule        = $ModuleName
            #FunctionsToExport = $FunctionNames
            Passthru          = $True
            Verbose           = $True
            ErrorAction       = "Stop"
        }
        If ($CurrentParams.CompanyName) {
            $Param_NewModule.Add("CompanyName",$CompanyName)
            $Param_NewModule.Copyright = "Copyright $CopyrightDate, $CompanyName"
        }
        Else {
            $Param_NewModule.Copyright = "Copyright $CopyrightDate, $Author"
        }

        $Msg = "Create new module manifest with the following specifications:`n$($Param_NewModule | Format-Table -AutoSize | Out-String)"
        $FGColor = "White"
        If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
        Else {Write-Verbose $Msg}

        $ConfirmMsg = "`n`n`tCreate module manifest for '$ModuleName'`n`n"
        If ($PSCmdlet.ShouldProcess($ModulePath,$ConfirmMsg)) {
        
            Try {
                New-ModuleManifest @Param_NewModule
                $Msg = "Successfully created new PowerShell module manifest for '$ModuleName' in '$ModulePath'"
                Write-Verbose $Msg
            }
            Catch {
                $Msg = "Failed to create PowerShell module manifest for '$ModuleName' in '$ModulePath'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
            }
        }
        Else {
            $Msg = "Operation cancelled by user"
            $FGColor = "White"
            If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
            Else {Write-Verbose $Msg}
        }
    } #end if not exist
}
End {}

} #end New-PKPSModuleManifest