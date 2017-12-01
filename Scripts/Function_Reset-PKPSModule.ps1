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
    Version : 01.02.0000
    Author  : Paula Kingsley
    
    History:

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK *

        v01.00.0000 - 2017-06-23 - Created script
        v01.01.0000 - 2017-09-06 - Added -passthru to Import-Module; minor cosmetic updates
        v02.00.0000 - 2017-12-01 - Overhauled, cosmetic updates, now imports module if not loaded

.EXAMPLE
    PS C:\> Reset-PKPSModule -Module pktools -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value           
        ---                   -----           
        Module                {pktools}       
        Verbose               True            
        SuppressConsoleOutput False           
        ScriptName            Reset-PKPSModule
        ScriptVersion         1.1.0           

        Action: Remove/import module(s)
        VERBOSE: Module name: pktools
        WARNING: Module 'pktools' is not currently loaded
        VERBOSE: Module imported successfully


        ComputerName : WORKSTATION14
        Name         : PKTools
        IsReset      : True
        Path         : C:\Users\jbloggs\scripts\PKTools\PKTools.psd1
        Type         : Script
        OldVersion   : -
        NewVersion   : 1.6.1
        Messages     : Module imported successfully        

.EXAMPLE
    PS C:\> Reset-PKPSModule -Module pktools

        Action: Remove/import module(s)
        WARNING: Loaded version of module 'PKTools' is identical to available version 1.6.1
        Module 'PKTools' removal/re-import cancelled by user


        ComputerName : WORKSTATION14
        Name         : PKTools
        IsReset      : False
        Path         : C:\Users\jbloggs\scripts\PKTools\PKTools.psd1
        Type         : Script
        OldVersion   : 1.6.1
        NewVersion   : -
        Messages     : Module 'PKTools' removal/re-import cancelled by user

.EXAMPLE
    PS C:\> $Arr | Reset-PKPSModule | FT -AutoSize

        Action: Remove/import module(s)
        WARNING: Module Sandbox is not currently loaded
        WARNING: Loaded version of module 'PKActiveDirectory' is identical to available version 2.7.0
        WARNING: Loaded version of module 'PKVMwareScripts' is identical to available version 3.13.0
        Module 'PKVMwareScripts' removal/re-import cancelled by user
        WARNING: Loaded version of module 'PKTools' is 1.7.0; available version is 1.6.2
        WARNING: Module 'foo' is not currently loaded
        ERROR: Module 'foo' not found in any module directory in path

        ComputerName  Name               IsReset Path                                                           Type   OldVersion NewVersion Messages                                              
        ------------  ----               ------- ----                                                           ----   ---------- ---------- --------                                              
        WORKSTATION19 Sandbox               True C:\Users\jbloggs\scripts\Sandbox\Sandbox.psm1                  Script -          1.0.0      Module removed/imported successfully                  
        WORKSTATION19 PKActiveDirectory     True C:\Users\jbloggs\corp\PKActiveDirectory\PKActiveDirectory.psm1 Script 2.7.0      2.7.0      Module removed/imported successfully                  
        WORKSTATION19 PKVMwareScripts       True C:\Users\jbloggs\corp\PKVmwareScripts\PKVmwareScripts.psm1     Script 3.13.0     -          Module 'PKVMwareScripts' removal/re-import cancelled by user
        WORKSTATION19 PKTools               True C:\Users\jbloggs\scripts\PKTools\PKTools.psm1                  Script 1.6.2      1.7.0      Module removed/imported successfully                  
        WORKSTATION19 foo                  Error Error                                                          Error  Error      Error      Module 'foo' not found in any module directory in path
    
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
    [object[]]$Module,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Suppress all non-verbose, non-error console output"
    )]
    [switch]$SuppressConsoleOutput
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"

    #Pipeline
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("Module")) -and (-not $Module)

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

    #region splats

    # General-purpse splat
    $StdParams = @{
        Verbose = $False
        ErrorAction = "Stop"
    }

    # Splat for Write-Progress
    $Activity = "Remove/import module(s)"
    $Param_WP1 = @{}
    $Param_WP1 = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }
    $Param_WP2 = @{}
    $Param_WP2 = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }


    #endregion Splats

    # Output obuect
    $InitialValue = "Error"
    $OutputTemplate = New-Object PSObject -Property ([ordered]@{
        ComputerName = $env:COMPUTERNAME
        Name         = $InitialValue
        IsReset      = $InitialValue        
        Path         = $InitialValue
        Type         = $InitialValue
        OldVersion   = $InitialValue
        NewVersion   = $InitialValue
        Messages     = $InitialValue
    })
    $Results = @()

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}
Process {
    
    $Total1 = $Module.Count
    $Current1 = 0

    # Go through list of module s
    Foreach ($M in $Module) {
        
        # create outer object
        $Output1 = $OutputTemplate.PSObject.Copy()
       
        If ($M -is [psmoduleinfo]) {
            $Output1.Name = $M.Name
            $Msg = "Module object: $($M.Name)"
        }
        Else {
            $Output1.Name = $M
            $Msg = "Module name: $($M)"
        }

       Write-Verbose $Msg

       $Current1 ++
       $Param_WP1.Status = "$($Output1.Module)"
       $Param_WP1.PercentComplete = ($Current1/$Total1*100)
       
       [switch]$Continue = $False

       # Make sure it's a valid module object and see if it's currently loaded
        Try {
            $Msg = "Get module object"
            $Param_WP1.CurrentOperation = $Msg
            Write-Progress $Param_WP1

            [Switch]$IsLoaded = $False
            If (-not ([array]$ModuleObj = (Get-Module $M @StdParams | Add-Member -MemberType NoteProperty -Name "IsLoaded" -Value $True -PassThru -Force | Select -Property *))) {

                $Msg = "Module '$M' is not currently loaded"
                Write-Warning $Msg

                If (-not ([array]$ModuleObj = Get-Module $M -ListAvailable @StdParams | Add-Member -MemberType NoteProperty -Name "IsLoaded" -Value $False -PassThru -Force | Select -Property *)) {
                    $Msg = "Module '$($Output1.Name)' not found in any module directory in path"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    $Output1.Messages = $Msg
                    $Results += $Output1
                }
                Else {
                    $Continue = $True
                }
            }
            Else {
                $Continue = $True
                $IsLoaded = $True
            }
        }
        Catch {
            $Msg = "Module search failed for $($Output.Name)"
            $Host.UI.WriteErrorLine("ERROR: $Msg for $M")
            $Output1.Messages = $Msg
            $Results += $Output1
        }
        
        
        # If we got the object    
        If ($Continue.IsPresent) {
            
            $Continue = $False

            $Total2 = $ModuleObj.Count
            $Current2 = 0

            Foreach ($Obj in $ModuleObj) {
            
                $Current2 ++
                $Param_WP2.Status = $Obj.Name
                $Param_WP2.PercentComplete = ($Current2/$Total2*100)
                
                $Msg = "Look for available version"
                $Param_WP2.CurrentOperation = $Msg
                Write-Progress @Param_WP2

                # create inner object
                $Output2 = $OutputTemplate.PSObject.Copy()
                $Output2.Name = $Obj.Name
                $Output2.Path = $Obj.Path
                $Output2.Type = $Obj.ModuleType

                $AvailableVer = (Get-Module $Obj.Name -ListAvailable @StdParams).Version 

                # If it's loaded, check the available version against the current one, and prompt to remove current
                If ($Obj.IsLoaded -eq $True) {
                    
                    $Output2.OldVersion = $Obj.Version

                    If ($Obj.Version -eq $AvailableVer) {
                        $Msg = "Loaded version of module '$($Obj.Name)' is identical to available version $AvailableVer"
                        Write-Warning $Msg
                    }
                    Else {
                        $Msg = "Loaded version of module '$($Obj.Name)' is $($Obj.Version); available version is $AvailableVer"
                        Write-Warning $Msg
                    }

                    [switch]$Continue = $False
                    $Msg = "Remove module"
                    $Param_WP2.CurrentOperation = $Msg
                    Write-Progress @Param_WP2

                    If ($PSCmdlet.ShouldProcess($($Obj.Name),$Msg)) {
                        Try {
                            $Null = $Obj | Remove-Module -Force -Confirm:$False @StdParams
                            $Continue = $True
                        }
                        Catch {
                            $Msg = "Module removal failed for $($Obj.Name)"
                            If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                            $Host.UI.WriteErrorLine("ERROR: $Msg")
                            $Output2.IsReset = $False
                            $Output2.Messages = $ErrorDetails
                        }
                    }
                    Else {
                        $Msg = "Module '$($Obj.Name)' removal/re-import cancelled by user"
                        $Host.UI.WriteErrorLine($Msg)
                        $Output2.NewVersion = "-"
                        $Output2.Messages = $Msg
                        $Output2.IsReset = $False
                    }
                
                } #end if it was loaded
                Else {
                    $Output2.OldVersion = "-"
                    $Continue = $True
                }

                # Import module
                If ($Continue.IsPresent) {
                    $Msg = "Import module"
                    $Param_WP2.CurrentOperation = $Msg
                    Write-Progress @Param_WP2

                    If ($PSCmdlet.ShouldProcess($Obj.Name,$Msg)) {
                        Try {
                        $Import = Get-Module $Obj.Name -ListAvailable @StdParams | Import-Module -PassThru -Force -Global @StdParams
                        
                        If ($Obj.IsLoaded -eq $True) {
                            $Msg = "Module removed/imported successfully"
                        }
                        Else {
                            $Msg = "Module imported successfully"
                        }
                        Write-Verbose $Msg
                        $Output2.IsReset = $True
                        $Output2.NewVersion = $Import.Version
                        $Output2.Messages = $Msg
                        }
                        Catch {
                            $Msg = "Module import failed for $($Obj.Name)"
                            If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                            $Host.UI.WriteErrorLine("ERROR: $Msg")
                            $Output2.Messages = $ErrorDetails
                            $Output2.IsReset = $False
                        }
                    }
                    Else {
                        $Msg = "Import of module '$($Obj.Name)' cancelled by user"
                        $Host.UI.WriteErrorLine($Msg)
                        $Output2.IsReset = $False
                        $Output2.NewVersion = "-"
                        $Output2.Messages = $Msg
                    }
                }

                # Add them up
                $Results += $Output2
            } #end for each found module object

        } #end for each valid module object
    
    } #end fo each input
        
    
}
End {
    Write-Output $Results
    $Null = Write-Progress -Activity $Activity -Completed -ErrorAction SilentlyContinue -Verbose:$False
}
} #end Reset-PKPSModule