#requires -version 3
Function Reset-PKPSModule {
<# 
.Synopsis
    Removes and re-imports named PowerShell modules   
    
.DESCRIPTION
   Removes and re-imports PowerShell modules   
   Prompts for confirmation
   Returns a PSObject 
   
.NOTES
    Name    : Function_Reset-PKPSModule.ps1
    Created : 2017-06-23
    Version : 01.01.0000
    Author  : Paula Kingsley
    
    History:

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK *

        v01.00.0000 - 2017-06-23 - Created script
        v01.01.0000 - 2017-09-06 - Added -passthru to Import-Module; minor cosmetic updates
        
.EXAMPLE
    PS C:\> Get-Module gn* | Reset-PKPSModule -Verbose

        VERBOSE: Remove and re-import module(s)
        VERBOSE: GNOpsActiveDirectory
        WARNING: Current version of module 'GNOpsActiveDirectory' is identical to available module version 2.4.11
        VERBOSE: GNOpsTools
        WARNING: Current version of module 'GNOpsTools' is identical to available module version 3.9.14
        VERBOSE: GNOpsWindowsChef
        WARNING: Current version of module 'GNOpsWindowsChef' is identical to available module version 2.5.0
        Removal of module 'GNOpsWindowsChef' cancelled by user
        VERBOSE: GNOpsWindowsVM
        WARNING: Current version of module 'GNOpsWindowsVM' is identical to available module version 5.1.6


        Module     : GNOpsActiveDirectory
        Path       : C:\Users\jbloggs\Git\Gracenote\Infrastructure\PowerShell\GNOpsActiveDirectory\GNOpsActiveDirectory.psm1
        Type       : Script
        OldVersion : 2.4.11
        NewVersion : 2.4.11
        Messages   : 

        Module     : GNOpsTools
        Path       : C:\Users\jbloggs\Git\Gracenote\Infrastructure\PowerShell\GNOpsTools\GNOpsTools.psm1
        Type       : Script
        OldVersion : 3.9.14
        NewVersion : 3.9.14
        Messages   : 

        Module     : GNOpsWindowsChef
        Path       : C:\Users\jbloggs\Git\Gracenote\Infrastructure\PowerShell\GNOpsWindowsChef\GNOpsWindowsChef.psm1
        Type       : Script
        OldVersion : 2.5.0
        NewVersion : 
        Messages   : Removal of module 'GNOpsWindowsChef' cancelled by user

        Module     : GNOpsWindowsVM
        Path       : C:\Users\jbloggs\Git\Gracenote\Infrastructure\PowerShell\GNOpsWindowsVM\GNOpsWindowsVM.psm1
        Type       : Script
        OldVersion : 5.1.6
        NewVersion : 5.1.6
        Messages   : 

    
#>

[CmdletBinding(
    SupportsShouldProcess=$True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Name of module, or module object"
    )]
    [alias("Name")]
    [object[]]$Module
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"

    # Generalpurpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"


    $ErrorActionPreference = "Stop"
    $StdParams = @{
        Verbose = $False
        ErrorAction = "Stop"
    }

    $OutputTemplate = New-Object PSObject -Property ([ordered]@{
        Module     = $M
        Path       = "Error"
        Type       = "Error"
        OldVersion = "Error"
        NewVersion = "Error"
        Messages   = "Error"
    })

    $Results = @()

    $Activity = "Remove/import module(s)"
    $Msg = $Activity
    Write-Verbose $Msg

}
Process {
    
    Foreach ($M in $Module) {
       
       $Output = $OutputTemplate.PSObject.Copy()
       $Output.Module = $M

        Try {
            
            [switch]$IsLoaded = $True
            If (-not ($ModuleObj = Get-Module $M @StdParams)) {

                $Msg = "Module '$M' is not currently loaded"
                Write-Warning $Msg
                [switch]$IsLoaded = $False

                If (-not ($ModuleObj = Get-Module $M -ListAvailable @StdParams)) {
                    $Msg = "Module '$M' not found in any module directory in path"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    Return
                }
            }

            If ($ModuleObj) {
                $Total = (($ModuleObj -as [array]).Count)
                $Current = 0

                Foreach ($Obj in $ModuleObj) {
                
                    $Current ++
                    Write-Verbose $Obj
                
                    $Output = $OutputTemplate.PSObject.Copy()
                    $AvailableVer = (Get-Module $Obj.Name -ListAvailable @StdParams).Version 

                    If ($IsLoaded.IsPresent) {
                        $Output.OldVersion = $Obj.Version
                        If ($Obj.Version -eq $AvailableVer) {
                            $Msg = "Loaded version of module '$($Obj.Name)' is identical to available version $AvailableVer"
                            Write-Warning $Msg
                        }
                        If (($Obj.Version -gt $AvailableVer) -or ($Obj.Version -lt $AvailableVer)) {
                            $Msg = "Loaded version of module '$($Obj.Name)' is $($Obj.Version); available version is $AvailableVer"
                            Write-Verbose $Msg
                        }
                    }
                    Else {$Output.OldVersion = "n/a"}

                    $Output.Module     = $Obj.Name
                    $Output.Path       = $Obj.Path
                    $Output.Type       = $Obj.ModuleType
                    $Output.NewVersion = "Error"
                    $Output.Messages   = "Error"
                
                    Write-Progress -Activity $Activity -CurrentOperation $Obj.Name -PercentComplete ($Current / $Total * 100)

                    If ($IsLoaded.IsPresent) {
                        [switch]$Continue = $False
                        $Msg = "Remove module"
                        If ($PSCmdlet.ShouldProcess($($Obj.Name),$Msg)) {
                            Try {
                                $Null = $Obj | Remove-Module -Force -Confirm:$False @StdParams
                                $Continue = $True
                            }
                            Catch {
                                $Msg = "Module $($Obj.Name)' removal failed"
                                $ErrorDetails = $_.Exception.Message
                                $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
                                $Output.Messages = $ErrorDetails
                                [switch]$Continue = $False
                            }
                        }
                        Else {
                            $Msg = "Module '$($Obj.Name)' reset cancelled by user"
                            Write-Warning $Msg
                            $Output.Messages = $Msg
                            [switch]$Continue = $False
                        }
                    }
                    Else {$Continue = $True}
                    
                    If ($Continue.IsPresent) {
                        $Msg = "Import module"
                        If ($PSCmdlet.ShouldProcess($Obj.Name,$Msg)) {
                            $Import = Get-Module $Obj.Name -ListAvailable @StdParams | Import-Module -PassThru -Force -Global @StdParams
                            $Output.NewVersion = $Import.Version
                            $Output.Messages = $Null
                        }
                        Else {
                            $Msg = "Import of module '$($Obj.Name)' cancelled by user"
                            Write-Verbose $Msg
                            $Output.NewVersion = $Null
                            $Output.Messages = $Msg
                        }
                    }

                    $Results += $Output 
                } #end for each module found
            } # end if module
        }
        Catch {
            $Msg = "Module lookup failed for '$M'"
            $ErrorDetails = $_.Exception.Message
            $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
            $Output.Messages = $ErrorDetails

            $Results += $Output        
        }
        
        
    } # end for each module
    
}
End {
    Write-Output $Results
    $Null = Write-Progress -Activity $Activity -Completed -ErrorAction SilentlyContinue -Verbose:$False
}
} #end Reset-PKPSModule