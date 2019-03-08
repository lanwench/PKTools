Function Save-PKPowerCLIModule {
<#
.SYNOPSIS
    Downloads and saves the PowerCLI VMware module from the PowerShell Gallery, with selection for version and path

.DESCRIPTION
    Downloads and saves the PowerCLI VMware module from the PowerShell Gallery, with selection for version and path
    Requires PowerShell 5
    Returns a string
    SupportsShouldProcess

.NOTES
    Name    : Function_Save-PKPowerCLIModule.ps1
    Author  : Paula Kingsley
    Created : 2019-02-26
    Version : 01.00.0000
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2019-02-26 - Created script based on Dimitar_Milov's original

.LINK
    https://www.powershellgallery.com/packages/VMware.PowerCLI/11.1.0.11289667

#>
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
param(
    [Parameter(
        #Mandatory = $true,
        HelpMessage = "Desired PowerCLI version (default is latest version)"
    )]
    [ValidateNotNullOrEmpty()]
    [version]$RequiredVersion,

    [Parameter(
        HelpMessage = "Absolute path for module on local computer"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript( { Test-Path $_} )]
    [string]$Path = "$Env:SystemDrive\Program Files\WindowsPowerShell\Modules",

    [Parameter(
        HelpMessage = "Repository name for PSGet (default PSGallery)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Repository = 'PSGallery',

    [Parameter(
        HelpMessage = "Force download/save even if version already exists in path"
    )]
    [switch]$Force,

    [Parameter(
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    #region Module info
    
    # Module name
    $PowerCLIModuleName = 'VMware.PowerCLI'

    # Dependency order
    $DependencyOrder = 'VMware.VimAutomation.Sdk', 'VMware.VimAutomation.Common', 'VMware.Vim', 'VMware.VimAutomation.Cis.Core', 'VMware.VimAutomation.Core', 'VMware.VimAutomation.Nsxt', 'VMware.VimAutomation.Vmc', 'VMware.VimAutomation.Vds', 'VMware.VimAutomation.Srm', 'VMware.ImageBuilder', 'VMware.VimAutomation.Storage', 'VMware.VimAutomation.StorageUtility', 'VMware.VimAutomation.License', 'VMware.VumAutomation', 'VMware.VimAutomation.HorizonView', 'VMware.DeployAutomation', 'VMware.VimAutomation.vROps', 'VMware.VimAutomation.PCloud'

    #endregion Module info

    #region Splats

    # Splat for Find-Module
    $Param_FindModule = @{}
    $Param_FindModule = @{
        Name        = $PowerCLIModuleName
        ErrorAction = "SilentlyContinue"
        Verbose     = $False
    }
    If ($CurrentParams.RequiredVersion) {
        $Param_FindModule.Add("RequiredVersion",$RequiredVersion)
    }

    # Splat for Save-Module
    $Param_SaveModule = @{}
    $Param_SaveModule = @{
        Path        = $Path
        Confirm     = $False
        Force       = $True
        ErrorAction = "Stop"
        Verbose     = $False
    }
    
    # Splat for write-progress
    $Activity = "Find and download PowerCLI from '$Repository'"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        Status           = $Null
        CurrentOperation = $Null
        ErrorAction      = "SilentlyContinue"
        Verbose          = $False
    }

    # Splat for remove-item
    $Param_Remove   = @{}
    $Param_Remove   = @{
        Confirm     = $False
        Force       = $True
        Recurse     = $True
        ErrorAction = "SilentlyContinue"                    
    }

    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "START   : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}

}
Process {
    
    [switch]$Continue = $False

    Try {
        $Msg = "Find module in repository"
        Write-Verbose $Msg
        $Param_WP.CurrentOperation = $Msg
        $Param_WP.Status = "$PowerCLIModuleName"
        Write-Progress @Param_WP

        If ($PowerCLIModuleObj = Find-Module @Param_FindModule) {
            $Msg = "Found PowerCLI module '$($PowerCLIModuleObj.Name)' version $($PowerCLIModuleObj.Version.ToString()) in repository '$($PowerCLIModuleObj.Repository)'"
            If (-not $CurrentParams.RequiredVersion) {$RequiredVersion = $PowerCLIModuleObj.Version}
            Write-Verbose $Msg
            $Continue = $True
        }
        Else {
            $Msg = "Failed to find PowerCLI module"
            If ($CurrentParams.RequiredVersion) {
                $Msg += " version $RequiredVersion"
            }
            $Msg += " in repository '$($CurrentParams.Repository)'"
            $Host.UI.WriteErrorLine("ERROR  : $Msg")
        }
    }
    Catch {
        $Msg = "Failed to find PowerCLI module"
        If ($CurrentParams.RequiredVersion) {
            $Msg += " version $RequiredVersion"
        }
        $Msg += " in repository '$($CurrentParams.Repository)'"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR  : $Msg")
    }

    # Check for existing version
    If ($Continue.IsPresent) {
        
        # Reset flag
        $Continue = $False

        # Update status
        $Msg = "Test for existing module files"
        Write-Verbose $Msg
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP

        # Make sure it isn't already there
        Try {
            If ($Exists = Get-Module -Name "$Path\$PowerCLIModuleName" -ListAvailable -ErrorAction SilentlyContinue) {
                If ($Exists.Version -eq $RequiredVersion) {
                    $Msg = "Module '$($Exists.Name)' version '$($Exists.Version)' already exists in path"
                    If ($Force.IsPresent) {
                        Write-Warning "$Msg; -Force detected"
                        $Continue = $True
                    }
                    Else {    
                        $Host.UI.WriteErrorLine("ERROR  : $Msg; specify -Force to overwrite")
                    }
                }
                Else {
                    $Msg = "Module '$($Exists.Name)' version '$($Exists.Version)' already exists in path; will be overwritten if you confirm next step" 
                    Write-Warning $Msg
                    $Continue = $True
                }
            }
            Else {
                $Msg = "Confirmed no matching module already exists in path"
                Write-Verbose $Msg
                $Continue = $True
            }
        }
        Catch {}
    }

    If ($Continue.IsPresent) {
        
        # Reset flag
        $Continue = $False
        
        # Update status
        $Msg = "Save module to path"
        Write-Verbose $Msg
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP
       
        $ConfirmMsg = "`n`n`tSave module '$($PowerCLIModuleObj.Name)' version $($PowerCLIModuleObj.Version.ToString())`n`tfrom repository '$Repository'`n`tto path '$Path'`n`n"
        If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$ConfirmMsg)) {
            Try {
                $Save = $PowerCLIModuleObj | Save-Module @Param_SaveModule 
                $Msg = "Ssaved module '$($PowerCLIModuleObj.Name)' to '$Path'"
                Write-Verbose $Msg
                
                # Reset flag
                $Continue = $True
            }
            Catch {
                $Msg = "Failed to save module '$($PowerCLIModuleObj.Name)' to path"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR  : $Msg")
            }
        }
        Else {
            $Msg = "Operation cancelled by user"
            Write-Warning $Msg
        }
    }

    # Save dependency modules (won't prompt for confirmation as we already have a Y or N above)
    If ($Continue.IsPresent) {
        
        # Reset flag
        $Continue = $False

        $Msg = "Save dependency modules to path"
        Write-Verbose $Msg
        $Param_WP.CurrentOperation = "Save dependency module"
        
        # Set the dependency order?
        $OrderedDependencies = @()
        foreach ($depModuleName in $DependencyOrder) {
            $OrderedDependencies += $PowerCLIModuleObj.Dependencies | Where-Object {$_.Name -eq $depModuleName}
        }

        # Save dependencies with minimum version
        $Count = 0
        $Total = $OrderedDependencies.Count
        Foreach ($Dependency in $OrderedDependencies) {
                                
            Try {
                $Param_WP.Status = $Dependency.Name
                Write-Progress @Param_WP

                $Param_FindModule.Name = $Dependency
                $Param_FindModule.RequiredVersion = $Dependency.MinimumVersion
                $Param_SaveModule.ErrorAction = "SilentlyContinue"

                $Save = (Find-Module @Param_FindModule | Save-Module @Param_SaveModule)
                $Count ++
                $Msg = "Saved dependency module '$($Dependency.Name)' to path '$Path'"
                Write-Verbose $Msg
            }
            Catch {
                $Msg = "Failed to save dependency module '$($Dependency.Name)' to path"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR  : $Msg")
            }
        } #end foreach
 
        If ($Count -eq $Total) {
            $Continue = $True
        }
        Else {
            $Msg = "At least one dependency module failed to save"
            Write-Warning $Msg
            $Continue = $True
        }
    }
    
    <#
    If ($Continue.IsPresent) {
            
        $Msg = "Remove non-matching dependency module versions from path"
        Write-Verbose $Msg
        $Param_WP.Status = $PowerCLIModuleObj.Name
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP
    
        # Remove newer dependencies versoin
        Foreach ($Dependency in $OrderedDependencies) {
            
            $ConfirmMsg = "`n`n`tRemove dependency module '$($Dependency.Name)' not matching minimum version`n`n"
            If ($PSCmdlet.ShouldProcess($ComputerName,$ConfirmMsg)) {
                Try {
                    $Remove = Get-ChildItem -Path (Join-Path $path $dependency.Name) -ErrorAction SilentlyContinue |
                        Where-Object {$_.Name -ne $dependency.MinimumVersion} | Remove-Item @Param_Remove
                    $Msg = "Removed non-matching version of '$($Dependency.Name)' from path"
                }
                Catch {
                    $Msg = "Failed to remove non-matching dependency module '$($Dependency.Name)' from path"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR  : $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                Write-Warning $Msg
            }
        } #end foreach        
    }
    #>

    # Make sure we're good
    If ($Continue.IsPresent) {
        Try {
            If ($Success = Get-Module -Name "$Path\$PowerCLIModuleName" -ListAvailable -ErrorAction SilentlyContinue) {
                If ($Success.Version -eq $RequiredVersion) {
                    $Msg = "Operation complete; please import module from path '$Path'"
                    $FGColor = "Green"
                    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
                    Else {Write-Verbose $Msg}
                }
                Else {
                    $Msg = "Failed to save module '$PowerCLI' version $RequiredVersion to path '$Path'" 
                    $Host.UI.WriteErrorLine("ERROR  : $Msg")
                }
            }
            Else {
                $Msg = "Failed to save module '$PowerCLI' version $RequiredVersion to path '$Path'" 
                $Host.UI.WriteErrorLine("ERROR  : $Msg")
            }
        }
        Catch {}
    }


}
} #end Save-PKPowerCLIModule