#Requires -version 3
Function Test-PKWSMan {
<# 
.SYNOPSIS
    Test WinRM connectivity to a remote computer using various protocols

.DESCRIPTION
    Test WinRM connectivity to a remote computer using various protocols
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Test-PKWSMan.ps1
    Created : 2018-10-23
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-10-23 - Created script
        v01.01.0000 - 2019-03-26 - Minor updates

.PARAMETER ComputerName
    One or more computer names

.PARAMETER Credential
    Valid credentials on target (default is current user credentials)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Test-PKWSMan

        
#> 

[CmdletBinding(
    DefaultParameterSetName = "__DefaultParameterSet",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

    [Parameter(
        HelpMessage="Valid credentials on target"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage="Authentication protocol: Basic, CredSSP, ClientCertificate, Digest, Kerberos,Negotiate, None (default Negotiate)"
    )]
    [ValidateSet("Basic","CredSSP","ClientCertificate","Digest","Kerberos","Negotiate","None")]
    [string] $Authentication = "Negotiate",

    [Parameter(
        HelpMessage = "Attempt to look up hostname in DNS if FQDN not provided"
    )]
    [Switch] $LookupDNS,

    [Parameter(
        HelpMessage = "Output mode: Full or Boolean (default is Full)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Full","Boolean")]
    [String] $OutputMode = "Full",

    [Parameter(
        HelpMessage ="Hide all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"

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
    
    # Output
    [array]$Results = @()
    
    #region Functions

    Function Test-DNS($Hostname) {
        If (-not ($Target = [System.Net.Dns]::GetHostByName($Hostname).Hostname)) {
            $Msg = "[$Hostname] Failed to resolve name in DNS using current nameservers and search paths"
            Write-Warning $Msg
        }
        Else {
            Write-Output $Target
        }
    }
    
    # Not currently using
    Function Test-Ping{
        Param($Computer)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
    }

    #endregion Functions

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Test-DNS
    $Param_DNS = @{}
    $Param_DNS = @{
        Hostname      = $Null
        ErrorAction   = "Silentlycontinue"
        WarningAction = $WarningPreference
    }

    # Splat for Test-WSMan
    $Param_WSMAN = @{}
    $Param_WSMAN = @{
        ComputerName   = $Null
        Authentication = $Authentication
        ErrorAction    = "Silentlycontinue"
        Verbose        = $VerbosePreference
        ErrorVariable  = "Nope"
    }

    # Splat for Write-Progress
    $Activity = "Test WSMAN connectivity"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg`n")}
    Else {Write-Verbose $Msg}


} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        $Computer = $Computer.Trim()
        $Current ++ 

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = $Computer
        
        [switch]$Continue = $False

        If ($Computer -in @($env:COMPUTERNAME,"LocalHost")) {
            $Msg = "Function cannot test WinRM connection to local computer"
            Write-Warning "[$Computer] $Msg"
        } 

        Else {
            $Msg = "Test WinRM connection using $Authentication authentication"
            
            $FGColor = "White"
            If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Computer] $Msg")}
            Else {Write-Verbose "[$Computer] $Msg"}

            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP
            
            $ConfirmMsg = "`n`n$Msg`n"
            If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
                $Start = Get-Date
                Try {
                    
                    # If this is hostname only
                    If ($Computer -notmatch "\w\.\w") {

                        If ($LookupDNS.IsPresent) {
                            $Msg = "Resolve hostname to FQDN in DNS"
                            $FGColor = "White"
                            If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Computer] $Msg")}
                            Else {Write-Verbose "[$Computer] $Msg"}

                            Try {
                                $Param_DNS.Hostname = $Computer
                                If ($Target = Test-DNS @Param_TestDNS) {
                                    
                                    $Computer = $Target

                                    $Msg = "Resolved hostname to FQDN"
                                    $FGColor = "White"
                                    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Computer] $Msg")}
                                    Else {Write-Verbose "[$Computer] $Msg"}
                                }
                                Else {
                                    $Msg = "Failed to resolve hostname to FQDN"
                                    Write-Warning "[$Computer] $Msg"
                                }
                            }
                            Catch {}
                        }
                        Else {
                            $Msg = "Target computer name does not appear to be a FQDN; connection may fail (try -LookupDNS)"
                            Write-Warning "[$Computer] $Msg"
                        }
                    } #end if not FQDN
                    
                    $Param_WSMAN.ComputerName = $Computer
                    $Test = Test-WSMAN @Param_WSMAN
                    $End = Get-Date

                    If ($Test) {
                        $Msg = "Successfully connected using $Authentication authentication in $($($End - $Start).Milliseconds) millisecond(s)"
                        If ($OutputMode -eq "Boolean") {
                            Write-Verbose "[$Computer] $Msg"
                            $True
                        }
                        Else {
                            $FGColor = "Green"
                            If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$Computer] $Msg")}
                            Else {Write-Warning "[$Computer] $Msg"}
                            #Write-Output "[$Computer] $Msg"
                        }
                        
                        
                    }
                    Else {
                        $Msg = "Failed to connect using $Authentication authentication after $($($End - $Start).Milliseconds) millisecond(s)"
                        If ($Nope) {
                            $Msg += "`n$([regex]:: match($Nope,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim())"
                        }
                        If ($OutputMode -eq "Boolean") {
                            Write-Warning "[$Computer] $Msg"
                            $False
                        }
                        Else {
                            $Host.UI.WriteErrorLine("ERROR  : [$Computer] $Msg")
                        }
                    }
                }
                Catch {}
            }
            Else {
                $Msg = "WinRM connection test cancelled by user"
                Write-Warning "[$Computer] $Msg"
            }
        }
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}
}

} # end Do-SomethingCool


$Null = New-Alias -Name Test-PKWinRM -Value Test-PKWSMan -Confirm:$False -Force -Verbose:$False