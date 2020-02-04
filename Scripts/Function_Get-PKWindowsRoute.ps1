#Requires -version 3
Function Get-PKWindowsRoute {
<#
.SYNOPSIS
    Invokes a scriptblock to get network routes on one or more computers using Get-NetRoute (available in PowerShell 4 on Windows 8/2012 and newer)

.DESCRIPTION
    Invokes a scriptblock to get network routes on one or more computers using Get-NetRoute (available in PowerShell 4 on Windows 8/2012 and newer)
    Accepts pipeline input
    Provides results on for 0.0.0.0/0, unless -AllDestinationPrefixes is present
    Optionally provides network adapter information where available
    Optionally tests connectivity to remote computers before invoking scriptblock
    Supports ShouldProcess
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsRoute.ps1
    Created : 2020-01-22
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-01-22 - Created script

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER AllDestinations
    Return results for all destination prefixes (default is 0.0.0.0/0)

.PARAMETER ExtendedAdapterDetails
    Return network adapter details

.PARAMETER Credential
    Valid credentials on target (default is passthrough)

.PARAMETER Authentication
    WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'Route')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE




#>


[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param (
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more computer names (default is local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Identity","Computer","Name","HostName","FQDN","DNSHostname")]
    [object[]]$ComputerName,
    
    [Parameter(
        HelpMessage = "Return results for all destination prefixes (default is 0.0.0.0/0)"
    )]
    [Switch] $AllDestinations,

    [Parameter(
        HelpMessage = "Return network adapter details NetConnectionID, MACAddress, and Name, if available"
    )]
    [Switch] $ExtendedAdapterDetails,

    [Parameter(
        HelpMessage = "Valid credentials on target (default is passthrough)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)"
    )]
    [ValidateSet('Kerberos','Basic','Negotiate','Default','CredSSP')]
    [string]$Authentication = "Negotiate",

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Run Invoke-Command scriptblock as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Prefix for job name (default is 'Route')"
    )]
    [String] $JobPrefix = "Route",

    [Parameter(
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

    [Parameter(
        HelpMessage = "Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    #$Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName
    }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        Param($AllDestinations,$ExtendedAdapterDetails)

        If ($ExtendedAdapterDetails.IsPresent) {
            
            $AdapterLookup = Get-WMIObject win32_networkadapter -Property DeviceID,Name,MACAddress,NetConnectionID | 
                Where-Object {$_.DeviceID -in (Get-WMIObject -Class Win32_NetworkAdapterConfiguration).Index} | 
                    Select-Object DeviceID,Name,MACAddress,NetConnectionID

            Function GetDeviceName($Index) {
                If (-not ($Results = ($AdapterLookup | Where-Object {$_.DeviceID -eq $Index}).Name)) {
                    $Results = "-"
                }
                $Results
            }
            Function GetName($Index) {
                If (-not ($Results = ($AdapterLookup | Where-Object {$_.DeviceID -eq $Index}).NetConnectionID)) {
                    $Results = "-"
                }
                $Results
                }
            Function GetMAC($Index) {
                If (-not ($Results = ($AdapterLookup | Where-Object {$_.DeviceID -eq $Index}).MACAddress)) {
                    $Results = "-"
                }
                $Results
            }
            
            $Select = @{N="ComputerName";E={$Env:ComputerName}},
                @{N="ifIndex";E={$_.ifIndex}},
                @{N="DestinationPrefix";E={$_.DestinationPrefix}},
                @{N="NextHop";E={$_.NextHop}},
                @{N="NetConnectionID";E={GetName -Index $_.ifIndex}},
                @{N="MACAddress";E={GetMac -Index $_.ifIndex}},
                @{N="Name";E={GetDeviceName -Index $_.ifIndex}},
                @{N="Messages";E={$Null}}

            $ErrorSelect = @{N="ComputerName";E={$Env:ComputerName}},
                @{N="ifIndex";E={"Error"}},
                @{N="DestinationPrefix";E={"Error"}},
                @{N="NextHop";E={"Error"}},
                @{N="NetConnectionID";E={"Error"}},
                @{N="MACAddress";E={"Error"}},
                @{N="Name";E={"Error"}},
                @{N="Messages";E={$Message}}
        }
        
        Else {
            $Select = @{N="ComputerName";E={$Env:ComputerName}},
                @{N="ifIndex";E={$_.ifIndex}},
                @{N="DestinationPrefix";E={$_.DestinationPrefix}},
                @{N="NextHop";E={$_.NextHop}},
                @{N="Messages";E={$Null}}
            $ErrorSelect = @{N="ComputerName";E={$Env:ComputerName}},
                @{N="ifIndex";E={"Error"}},
                @{N="DestinationPrefix";E={"Error"}},
                @{N="NextHop";E={"Error"}},
                @{N="Messages";E={$Message}}
        }


        If ($Null = Get-Command -Name Get-Netroute -ErrorAction SilentlyContinue) {
            
            $Param = @{}
            $Param = @{
                AddressFamily = "IPv4"
                ErrorAction   = "Stop"
                Verbose       = $False
            }
            If (-not $AllDestinations.IsPresent) {
                $Param.Add("DestinationPrefix","0.0.0.0/0")
            }
            
            Try {
                Get-NetRoute @Param |Sort-Object DestinationPrefix | Select-Object $Select
            }
            Catch {
                $Message = $_.Exception.Message
                "" | Select-Object $ErrorSelect
            }
        }
        Else {
            $Message = "Command requires PS4 and Windows 8/2012 at minimum"
            "" | Select-Object $ErrorSelect
        }

    } #end scriptblock

    #endregion Scriptblock for Invoke-Command

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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    # Function to test WinRM connectivity
    Function Test-WinRM{
        Param($Computer)
        $Param_WSMAN = @{
            ComputerName   = $Computer
            Credential     = $Credential
            Authentication = $Authentication
            ErrorAction    = "Silentlycontinue"
            Verbose        = $False
        }
        Try {
            If (Test-WSMan @Param_WSMAN) {$True}
            Else {$False}
        }
        Catch {$False}
    }

    # Function to test ping connectivity
    Function Test-Ping{
        Param($Computer)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
    }

    #endregion Functions

    #region Splats

    # Splat for Write-Progress
    $Activity = "Invoke scriptblock to return network routes"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity (as job)"
    }
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    # Parameters for Invoke-Command
    $ConfirmMsg = $Activity
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName   = $Null
        Authentication = $Authentication
        ArgumentList   = $AllDestinations,$ExtendedAdapterDetails
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.Add("AsJob",$True)
        $Param_IC.Add("JobName",$Null)
    }
    
    #endregion Splats

    # Console output
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title


} #end begin


Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        If ($Computer -is [string]) {
            $Computer = $Computer.Trim()
        }
        Elseif ($Computer -is [Microsoft.ActiveDirectory.Management.ADAccount]) {
            If ($Computer.DNSHostName) {
                $Computer = $Computer.DNSHostName
            }
            Else {
                $Computer = $Computer.Name
            }
        }
        
        $Current ++ 
        $Param_WP.PercentComplete = ($Current/$Total* 100)
        $Param_WP.Status = $Computer
        
        [switch]$Continue = $False

        Switch ($ConnectionTest) {
            Default {$Continue = $True}
            Ping {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Ping computer"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "Ping failure"
                            "[$Computer] $Msg" | Write-MessageError
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                    }
                }
                Else {$Continue = $True}
            }
            WinRM {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Test WinRM connection"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-WinRM -Computer $Computer) {
                            $Continue = $True
                        }
                        Else {
                            $Msg = "WinRM failure"
                            "[$Computer] $Msg" | Write-MessageError
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                    }
                }
                Else {$Continue = $True}
            }        
        }

        If ($Continue.IsPresent) {
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Msg = "Invoke command"
                    If ($AsJob.IsPresent) {$Msg += " as PSJob"}
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "evt_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Jobs += $Job
                    }
                    Else {
                        Invoke-Command @Param_IC | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[$Computer] $Msg" | Write-MessageError
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

    If ($AsJob.IsPresent -and ($Jobs.Count -gt 0)) {

        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output`n"
            "$Msg" | Write-MessageInfo -FGColor White -Title
            $Jobs | Get-Job
            
        }
        Else {
            $Msg = "No jobs created"
            $Msg | Write-MessageError
        }
    } #end if AsJob


    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title


}

} # end function

$Null = New-Alias -Name Get-PKWinRoute -Value Get-PKWindowsRoute -Confirm:$False -Force -Description "For guessability"