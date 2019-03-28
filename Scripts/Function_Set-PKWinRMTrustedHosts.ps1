#Requires -Version 3
Function Set-PKWinRMTrustedHosts {
<#
.SYNOPSIS
    Uses Set-Item to modify the trusted hosts for WinRM configured on the local computer

.DESCRIPTION
    Uses Get-Item to return the trusted hosts for WinRM configured on the local computer
    Returns a PSobject

.NOTES        
    Name    : Function_Set-PKWinRMTrustedHosts.ps1
    Created : 2019-04-18
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-03-18 - Created script

.LINK
    https://devblogs.microsoft.com/scripting/powertip-use-powershell-to-view-trusted-hosts/
    
.LINK
    https://www.dtonias.com/add-computers-trustedhosts-list-powershell/

.EXAMPLE
    PS C:\> Set-PKWinRMTrustedHosts -Value "*.internal.lan" -Verbose

        


#>
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        HelpMessage = "Action: Append, Clear, Overwrite, Remove value (default is 'Append')"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Append","Remove","Clear","Overwrite")]
    [String]$Action = "Append",

    [Parameter(
        HelpMessage = "Value to add to, or remove from, TrustedHosts list (e.g., '*.internal.domain.local')"
    )]
    [string]$Value,

    [Parameter(
        HelpMessage = "Force setting even if value already present"
    )]
    [Switch] $Force,

    [Parameter(
        HelpMessage = "Valid credentials on target"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential,# = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "Hide all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet
)

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region General variables

    # Make sure we have something to do
    If (-not $Value -and $Action -ne "Clear" ) {
        $Msg = "Parameter 'Value' is mandatory unless action is set to Clear"
        $Host.UI.WriteErrorLine("$Msg")
        Break
    }
    
    # Make sure we're admin
    If (-not([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")){
        $Msg = "Administrator rights are required to modify the TrustedHosts list"
        $Host.UI.WriteErrorLine("$Msg")
        Break
    }

    #Set path
    $Path = "WSMan:\localhost\Client\TrustedHosts"

    # Make sure we're being careful
    If ($Value -eq "*" -and ($Action -in "Append","Replace")) {
        $Msg = "Setting TrustedHosts to '*' is not recommended; please consider restricting the value to an internal domain such as '*.servers.domain.local'"
        Write-Warning $Msg
    }

    #endregion General variables

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Get-Item
    $Param_Get = @{}
    $Param_Get = @{
        Path          = $Path
        ErrorAction   = "Stop"
        WarningAction = $WarningPreference
    }
    If ($CurrentParams.Credential.Username) {
        $Param_Get.Add("Credential",$Credential)
    }
    
    # Splat for Get-ChildItem
    $Param_GCI = @{}
    $Param_GCI = @{
        Path          = $Path
        ErrorAction   = "Stop"
        WarningAction = $WarningPreference
    }
    
    # Splat for Set-Item
    $Param_Set = @{}
    $Param_Set = @{
        Value         = $Null
        Path          = $Path
        Force         = $True # This is separate from the function parameter; included here to avoid dual confirmation popup
        Passthru      = $True
        Confirm       = $False
        ErrorAction   = "SilentlyContinue"
        WarningAction = "Continue"
    }
    If ($CurrentParams.Credential.Username) {
        $Param_Set.Add("Credential",$Credential)
    }

    # Splat for Clear-Item
    $Param_Clear = @{}
    $Param_Clear = @{
        Path          = $Path
        Force         = $True # This is separate from the function parameter; included here to avoid dual confirmation popup
        Confirm       = $False
        ErrorAction   = "SilentlyContinue"
        WarningAction = "Continue"
    }
    If ($CurrentParams.Credential.Username) {
        $Param_Clear.Add("Credential",$Credential)
    }

    # Splat for Write-Progress
    Switch ($Action) {
        Append    {$Activity = "Append TrustedHosts value to WinRM client"}
        Clear     {$Activity = "Clear all TrustedHosts values on WinRM client"}
        Remove    {$Activity = "Remove TrustedHosts value from WinRM client"}
        Overwrite {$Activity = "Overwrite TrustedHosts value on WinRM client"}
    }
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = $Env:ComputerName
    }

    #endregion Splats

    #region Functions


    #endregion Functions

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOuptut.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}

}
Process {    
    

    [switch]$Continue = $False
    $Msg = "Get current TrustedHosts values"
    
    $FGColor = "White"
    If (-not $SuppressConsoleOuptut.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Env:ComputerName] $Msg")}
    Else {Write-Verbose "[$Env:ComputerName] $Msg"}
    
    $Param_WP.CurrentOperation = $Msg
    Write-Progress @Param_WP

    Try {
        
        $Current = (Get-Item @Param_Get | Select @{N="ComputerName";E={$env:COMPUTERNAME}},Name,SourceOfValue,Value)
        Switch ($Action) {
            Append {
                If ($Value -eq "*") {
                    $Msg  = "Value '$Value' is already in the TrustedHosts list"
                    If ($Force.IsPresent) {
                        $Continue = $True
                        $Msg += "; -Force specified"
                        $FGColor = "Cyan"
                    }
                    Else {
                        $Msg += "; -Force not specified"
                        $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
                        $Value = "$Value, $($Current.Value)"
                    }
                }
                ElseIf ($Current.Value -match [Regex]::Escape($Value)) {
                    $Msg  = "Value '$Value' is already in the TrustedHosts list"
                    If ($Force.IsPresent) {
                        $Continue = $True
                        $Msg += "; -Force specified"
                        $FGColor = "Cyan"
                    }
                    Else {
                        $Msg += "; -Force not specified"
                        $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
                        $Value = "$Value, $($Current.Value)"
                    }
                }
                Else {
                    $Msg  = "Value '$Value' does not exist in the TrustedHosts list"
                    If ($Current.Value) {$Value = "$Value, $($Current.Value)"}
                    $Continue = $True
                    $FGColor = "Green"
                }
            }
            Clear {
                If (-not $Current.Value) {
                    $Msg  = "TrustedHosts list is empty"
                    $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
                }
                Else {
                    $Msg  = "TrustedHosts list is not empty"
                    $FGColor = "Green"
                    $Continue = $True
                }
                
            }
            Remove {
                $Current
                If ($Current.Value -notmatch [Regex]::Escape($Value)) {
                    $Msg  = "Value '$Value' not found in the TrustedHosts list"
                    $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
                }
                Else {
                    $Msg  = "Value '$Value' found in the TrustedHosts list"
                    $FGColor = "Green"
                    $Continue = $True
                }
            }
            Overwrite {
                If ($Current.Value -ne $Value) {
                    $Msg = "TrustedHosts list will be overwritten"
                    $FGColor = "Green"
                    $Continue = $True
                }
                Else {
                    $Msg  = "TrustedHosts list is already set to '$Value'"
                    If ($Force.IsPresent) {
                        $Continue = $True
                        $Msg += "; -Force specified"
                        $FGColor = "Cyan"
                    }
                    Else {
                        $Msg += "; -Force not specified"
                        $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
                    }
                }
            }
            
        } #end switch
    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR  : [$Env:ComputerName] $Msg")
    }
            
    If ($Continue.IsPresent) {
        
        If (-not $SuppressConsoleOuptut.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Env:ComputerName] $Msg")}
        Else {Write-Verbose "[$Env:ComputerName] $Msg"}    

        $DefaultConfirmMsg = "WARNING: This command modifies the TrustedHosts list for the WinRM client.`n`t * The computers in the TrustedHosts list might not be authenticated.`n`t * The client might send credential information to these computers."
        $FGColor = "White"
        Switch ($Action) {
            
            Clear {
                $Msg = "Clear TrustedHosts list"
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
            }
            Remove {
                $Msg = "Remove '$Value' from TrustedHosts list"
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
            }
            Default {
                Switch ($Action) {
                    Append  {
                        $Msg = "Set TrustedHosts list to '$Value'"
                        $Param_WP.CurrentOperation = $Msg
                        Write-Progress @Param_WP

                    }
                    Overwrite {
                        $Msg = "Overwrite current TrustedHosts list '$($Current.Value)' with '$Value'"
                        $FGColor = "White"
                        $Param_WP.CurrentOperation = $Msg
                        Write-Progress @Param_WP
                        
                    }
                }
            }
        } #end switch

        # For confirmation prompt
        $ConfirmMsg = "`n`n`t$Msg`n`n`t$DefaultConfirmMsg`n`n"

        # Console output
        If (-not $SuppressConsoleOuptut.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Env:ComputerName] $Msg")}
        Else {Write-Verbose "[$Env:ComputerName] $Msg"}

        # Prompt to make change
        If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
            Try {   
                Switch ($Action) {
                    Clear {
                        $Clear = Clear-Item @Param_Clear
                        $New = (Get-Item @Param_Get).Value
                        If (-not $New) {
                            $FGColor = "Green"
                            $Msg = "Clear operation successful"
                        }
                        Else {
                            $Msg = "Clear operation failed"
                        }
                    }
                    Remove {

                        If (-not ($ToKeep = $Current.Value | Where-Object {$_ -notmatch [Regex]::Escape($Value)})) {
                            $ToKeep = " "
                        }
                        $Param_Set.Value = $ToKeep
                        $SetIt = Set-Item @Param_Set
                        $New = (Get-Item @Param_Get).Value
                        If ($New -notmatch [Regex]::Escape($Value)) {
                            $FGColor = "Green"
                            $Msg = "Remove operation successful"
                        }
                        Else {
                            $Msg = "Remove operation failed"
                        }

                        <#

                        If ($ToKeep = $Current.Value | Where-Object {$_ -notmatch [Regex]::Escape($Value)}) {
                            $Param_Set.Value = $ToKeep
                            $SetIt = Set-Item @Param_Set
                            $New = (Get-Item @Param_Get).Value
                            If ($New -notmatch [Regex]::Escape($Value)) {
                                $FGColor = "Green"
                                $Msg = "Remove operation successful"
                            }
                            Else {
                                $Msg = "Remove operation failed"
                            }
                        }
                        Else {
                            $Clear = Clear-Item @Param_Clear
                            $New = (Get-Item @Param_Get).Value
                            If (-not $New) {
                                $FGColor = "Green"
                                $Msg = "Remove operation successful"
                            }
                            Else {
                                $Msg = "Remove operation failed"
                            }
                        }

                        #>
                    }
                    Default {
                        
                        $Param_Set.Value = $Value
                        $SetIt = Set-Item @Param_Set
                        $New = (Get-Item @Param_Get).Value
                        
                        Switch ($Action) {
                            Append  {
                                If ($Value -in $New) {
                                    $FGColor = "Green"
                                    $Msg = "Append operation successful"
                                }
                                Else {
                                    $Msg = "Append operation failed"
                                }
                            }
                            Overwrite {
                                If ($New -eq $Value) {
                                    $FGColor = "Green"
                                    $Msg = "Overwrite operation successful"
                                }
                                Else {
                                    $Msg = "Overwrite operation failed"
                                }
                            }
                        } # end switch
                    } #end default
                } #end switch

                If (-not $SuppressConsoleOuptut.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Env:ComputerName] $Msg")}
                Else {Write-Verbose "[$Env:ComputerName] $Msg"}

                $Current = (Get-Item @Param_Get | Select @{N="ComputerName";E={$env:COMPUTERNAME}},Name,SourceOfValue,Value)
                Write-Output $Current

            } 
            Catch {
                $Msg = "$Action operation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
            }
        } #end confirm
        Else {
            $Msg = "Operation cancelled by user"
            $FGColor = "White"
            If (-not $SuppressConsoleOuptut.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Env:ComputerName] $Msg")}
            Else {Write-Verbose "[$Env:ComputerName] $Msg"}
        }    
    }
    
}
End {
    
    Write-Progress -Activity $Activity -Completed 

    #$Msg = "$($Current | Out-String)"
    #$FGColor = "White"
    #If (-not $SuppressConsoleOuptut.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    #Else {Write-Verbose "$Msg"}

    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOuptut.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}
}
} #end function

$null = New-Alias Set-TrustedHosts -Value Set-PKWinRMTrustedHosts -Description "Easier to remember!" -Force -Confirm:$False
$null = New-Alias Set-PKTrustedHosts -Value Set-PKWinRMTrustedHosts -Description "Easier to remember!" -Force -Confirm:$False