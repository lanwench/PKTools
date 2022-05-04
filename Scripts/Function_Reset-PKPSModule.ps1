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
    Version : 03.00.0000
    Author  : Paula Kingsley
    
    History:

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK *

        v01.00.0000 - 2017-06-23 - Created script
        v01.01.0000 - 2017-09-06 - Added -passthru to Import-Module; minor cosmetic updates
        v02.00.0000 - 2017-12-01 - Overhauled, cosmetic updates, now imports module if not loaded
        v03.00.0000 - 2022-04-26 - Total overhaul & simplificairton

.PARAMETER Module
    One or more module names or objects

.PARAMETER Force
    Import module even if current version matches

.EXAMPLE
    PS C:\> Reset-PKPSModule Kittens -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value           
        ---           -----           
        Verbose       True            
        Module        {Kittens}       
        Force         False           
        ScriptName    Reset-PKPSModule
        PipelineInput False           
        ScriptVersion 3.0.0           

        VERBOSE: [BEGIN: Reset-PKPSModule] Remove/import module(s)

        Status    Name    Version Path                                               
        ------    ----    ------- ----                                               
        Available Kittens 1.39.0 c:\repos\modules\Kittens\Kittens.psd1
        Imported  Kittens 1.38.0 c:\repos\modules\Kittens\Kittens.psm1

        VERBOSE: [Kittens] Available module version is 1.39.0; current module version is 1.38.0
        VERBOSE: [Kittens] Import module Kittens version 1.39.0 from c:\repos\modules\Kittens\Kittens.psd1

        ModuleType Version    Name       ExportedCommands                                                                                                                                                               
        ---------- -------    ----       ----------------                                                                                                                                                               
        Script     1.39.0     Kittens    {Set-LaserPointer, Open-FoodCan, Open-Box...}       

.EXAMPLE
    PS C:\> Reset-PKPSModule Puppies -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value           
        ---           -----           
        Verbose       True            
        Module        {Puppies}       
        Force         False           
        ScriptName    Reset-PKPSModule
        PipelineInput False           
        ScriptVersion 3.0.0           

        VERBOSE: [BEGIN: Reset-PKPSModule] Remove/import module(s)

        Status    Name    Version Path                                               
        ------    ----    ------- ----                                               
        Available Puppies 2.01.00 c:\repos\modules\Puppies\Puppies.psd1
        Imported  Puppies 2.01.00 c:\repos\modules\Puppies\Puppies.psm1

        VERBOSE: [Puppies] Available module version is 2.01.0; current module version is 2.01.0
        WARNING: [Puppies] Current module version is identical to available module version; -Force not specified

#>

[CmdletBinding(
    SupportsShouldProcess=$True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Position = 0,
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more module names or objects"
    )]
    [alias("Name")]
    [object[]]$Module,

    [Parameter(
        HelpMessage = "Import module even if current version matches"
    )]
    [switch]$Force
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "03.00.0000"

    # How did we get here?
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region splats

    # Splat for Write-Progress
    $Activity = "Remove/import module(s)"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    #endregion Splats
    
    $Msg = "[BEGIN: $Scriptname] $Activity" 
    Write-Verbose $Msg

}
Process {
    
    $Total = $Module.Count
    $Current = 0

    Foreach ($M in $Module) {
            
        If ($M -is [psmoduleinfo]) {$Name = $M.Name}
        Elseif ($M -is [string]) {$Name = $M}
        If ($Name) {
            
            $Output = @()
                
            [switch]$Continue = $False
            Write-Progress -Activity $Activity -CurrentOperation $Name -PercentComplete (($Current++)/$($Module.Count)*100)
            
            If ($Available = Get-Module $M -ListAvailable -ErrorAction SilentlyContinue -Verbose:$False) {
                
                $Output += $Available | Select @{N="Status";E={"Available"}},Name,Version,Path 
                $Msg = "Available module version is $($Available.Version.ToString())"
                
                If ($Imported = Get-Module $M -ErrorAction SilentlyContinue -Verbose:$False) {
                    
                    $Output += $Imported | Select @{N="Status";E={"Imported"}},Name,Version,Path
                    Write-Host ($Output | Format-Table -AutoSize | Out-String)
                    
                    $Msg += "; current module version is $($Imported.Version.ToString())"
                    Write-Verbose "[$Name] $Msg"

                    If (($Imported.Version.ToString() -lt $($Available.Version.ToString()))) {$Continue = $True}
                    ElseIf ($Force.IsPresent) {$Continue = $True}
                    Else {
                        $Msg = "Current module version is identical to available module version; -Force not specified"
                        Write-Warning "[$Name] $Msg"
                    }
                }
                Else {
                    Write-Verbose "[$Name] $Msg"
                    Write-Host ($Output | Format-Table -AutoSize | Out-String)
                }
                If (-not $Imported) {
                    $Msg = "Module is not currently imported"
                    Write-Warning "[$Name] $Msg"
                    $Continue = $True
                }
                         
                If ($Continue.IsPresent) {
                    $Msg = "Import module $($Available.Name) version $($Available.Version.ToString()) from $($Available.Path)"
                    Write-Verbose "[$Name] $Msg"
                    If ($PSCmdlet.ShouldProcess($Env:ComputerName,$Msg)) {
                        Try {
                            $Available | Import-Module -Force -PassThru -Global -Verbose:$False -ErrorAction Stop
                        }
                        Catch {
                            $Msg = "Failed to import module"
                            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                            Write-Warning "[$Name] $Msg"
                        }
                    }
                    Else {
                        $Msg = "Operation cancelled by user"
                        Write-Warning "[$Name] $Msg"
                    }
                }

            } #end if available

        } # end if name

    } #end for each module

}
End {
    $Null = Write-Progress -Activity $Activity -Completed -ErrorAction SilentlyContinue -Verbose:$False
}
} #end Reset-PKPSModule