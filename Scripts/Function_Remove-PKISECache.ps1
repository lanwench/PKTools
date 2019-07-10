#requires -version 3
Function Remove-PKISECache {
<#
.SYNOPSIS
    Removes PowerShell ISE cache files for the local user

.DESCRIPTION
    Removes PowerShell ISE cache files for the local user
    If files are found, stops any running PowerShell ISE processes before proceeding
    May help with slowness / 'intellisense timed out' messages
    Requires elevation and will not run within ISE, for obvious reasons
    Supports ShouldProcess

.NOTES
    Name    : Function_Remove-PKISECache.ps1 
    Created : 2019-06-27
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.00.0000 - 2019-06-27 - Created script

.LINK
    https://www.itprotoday.com/mobile-management-and-security/solve-intellisense-hangs-integrated-scripting-environment

.EXAMPLE
    PS C:\> Remove-PKISECache -Verbose

#>
[Cmdletbinding(
    SupportsShouldProcess,
    ConfirmImpact="High"
)]
Param(
    
    [Parameter(
        ValueFromPipeline = $True,
        HelpMessage = "Absolute path to local ISE cache directory (default is `$Env:LOCALAPPDATA\Microsoft\Windows\PowerShell\CommandAnalysis)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If (Test-Path $Path) {$True}})]
    $Path = "$Env:LOCALAPPDATA\Microsoft\Windows\PowerShell\CommandAnalysis"
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    # Display our parameters
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If ($PipelineInput.IsPresent) {$CurrentParams.Path = $Null}
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # Make sure we aren't trying to commit suicide
    If ($Host.Name -match "ISE") {
        $Msg = "You can't run this from inside the ISE, silly!"
        Write-Warning "[$Env:ComputerName] $Msg"
        Break
    }

    # Make sure we're elevated
    # https://superuser.com/questions/749243/detect-if-powershell-is-running-as-administrator
    If (-not ([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544"))) {
        $Msg = "PowerShell must be running in Elevated mode"
        Write-Warning "[$Env:ComputerName] $Msg"
        Break
    }

    
    #region Functions

    # Function to write a console message or a verbose message
    Function Write-MessageInfo {
        Param([Parameter(ValueFromPipeline)]$Message,$FGColor,[switch]$Title)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {
            If ($Title.IsPresent) {$Message = "`n$Message`n"}
            $Host.UI.WriteLine($FGColor,$BGColor,"$Message")
        }
        Else {Write-Verbose "$Message"}
    }

    # Function to write an error or a verbose message
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)#,[switch]$Quiet = $Quiet)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("$Message")}
        Else {Write-Verbose "$Message"}
    }

    #endregion Functions
    
    $Activity = "Clear PowerShell ISE cache files"
    $Msg = "BEGIN: $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}
Process {


    $Total = 3
    $Current = 0
    [switch]$Continue = $False


    # First let's make sure we have some
    $Msg = "Look for PowerShell ISE cache files"
    Write-Verbose "[$Env:ComputerName] $Msg"
    
    $Current ++
    Write-Progress -Activity $Activity -CurrentOperation $Msg -PercentComplete ($Current/$Total*100)

    Try {
        If ($CacheFiles = (Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue)) {
            $Msg = "$($CacheFiles) cache files found in '$Path'"
            Write-Verbose "[$Env:ComputerName] $Msg"
        
            $Msg = "Stop any running PowerShell ISE processes"
            Write-Verbose "[$Env:ComputerName] $Msg"
    
            $Current ++
            Write-Progress -Activity $Activity -CurrentOperation $Msg -PercentComplete ($Current/$Total*100)

            Try {

                If ($ISEProc = Get-Process -Name powershell_ise -ErrorAction Stop -Verbose:$False) {
            
                    $ConfirmMsg = "`n`n`tStop $($ISEProc.Count) running PowerShell ISE processes`n`n"
                    If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
                        $Msg = "Stopping $($ISEProc.Count) running PowerShell ISE processes"
                        Write-Verbose "[$Env:ComputerName] $Msg"
                        Try {
                            $ISEProc | Stop-Process -Force -Confirm:$False -ErrorAction Stop -Verbose:$False
                            $Msg = "Successfully stopped processes"
                            Write-Verbose "[$Env:ComputerName] $Msg"
                            $Continue = $True
                        }
                        Catch {
                            $Msg = "Failed to stop processes"
                            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails"}
                            Write-Error "[$Env:ComputerName] $Msg"
                            Break
                        }
                    }
                    Else {
                        $Msg = "Operation cancelled by user"
                        Write-Verbose "[$Env:ComputerName] $Msg"
                    }
                }
                Else {
                    $Msg = "No PowerShell ISE processes found"
                    Write-Verbose "[$Env:ComputerName] $Msg" 
                    $Continue = $True
                }

                If ($Continue.IsPresent) {
            
                    $Msg = "Remove PowerShell ISE cache files for user '$Env:UserName'"
                    Write-Verbose "[$Env:ComputerName] $Msg"

                    $Current ++
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -PercentComplete ($Current/$Total*100)

                    $Continue = $False
                    Try {
                        
                        $ConfirmMsg = "`n`n`tRemove $($CacheFiles.Count) cache files from`n`t$Path`n`n"
                        If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
                            $Msg = "Removing $($CacheFiles.Count) cache files from`n`t$Path"
                            Write-Verbose "[$Env:ComputerName] $Msg"
                            Try {
                                $CacheFiles | Remove-Item -Recurse -Force -Confirm:$False -ErrorAction Stop -Verbose:$False
                                $Msg = "Successfully removed cache files"
                                Write-Verbose "[$Env:ComputerName] $Msg"
                            
                            }
                            Catch {
                                $Msg = "Failed to remove cache files"
                                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails"}
                                Write-Error "[$Env:ComputerName] $Msg"
                                Break
                            }
                        }
                        Else {
                            $Msg = "Operation cancelled by user"
                            Write-Verbose "[$Env:ComputerName] $Msg"
                        }
                    }
                    Catch {
                        $Msg = "Failed to remove cache files"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails"}
                        Write-Error "[$Env:ComputerName] $Msg"
                    }
                } #end if continue
            }
            Catch {
                $Msg = "Failed to get PowerShell ISE processes"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails"}
                Write-Error $Msg
                Break
    
            }
        }
        Else {
            $Msg = "No cache files found in '$Path'"
            Write-Verbose "[$Env:ComputerName] $Msg"
        }
    }
    Catch {
        $Msg = "Failed to get cache files"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails"}
        Write-Error "[$Env:ComputerName] $Msg"
        Break
    }

}
End {
    
    Write-Progress -Activity $Activity -Completed
    If ($ISEProc) {
        $Msg = "You will need to restart PowerShell ISE manually"
        Write-Verbose "[$Env:ComputerName] $Msg"
    }
    
    $Msg = "END  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title
}
} #end Remove-PKISECache
