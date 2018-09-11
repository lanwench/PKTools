#Requires -version 3
Function Get-PKWindowsNetConfig {
<# 
.SYNOPSIS
    Gets Windows IPv4 network configuration on a local or remote computer, interactively or as a PSJob

.DESCRIPTION
    Gets Windows IPv4 network configuration on a local or remote computer, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsIPv4Config.ps1
    Created : 2018-08-29
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-08-29 - Created script
        

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER Credential
    Valid credentials on target computer (default is current user credentials)

.PARAMETER AsJob
    Invoke the command as a PS job

.PARAMETER ConnectionTest
    Run WinRM or ping test prior to invoke-command, or no test (default is WinRM)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKWindowsNetConfig -Verbose 

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Verbose               True                                     
        ComputerName          {WORKSTATION-14}                        
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        ConnectionTest        WinRM                                    
        SuppressConsoleOutput False                                    
        ScriptName            Get-PKWindowsNetConfig                   
        ScriptVersion         1.0.0                                    


        Action: Get Windows network configuration

        VERBOSE: [WORKSTATION-14] Invoke command

        ComputerName         : WORKSTATION-14
        Name                 : Ethernet
        InterfaceDescription : Cisco AnyConnect Secure Mobility Client Virtual Miniport Adapter for Windows x64
        InterfaceIndex       : 13
        MACAddress           : 00-05-9A-3C-7A-00
        Status               : Not Present
        DHCPEnabled          : False
        IPAddress            : 
        DefaultGateway       : 
        SubnetMask           : 
        DNSServers           : 
        Messages             : 

        ComputerName         : WORKSTATION-14
        Name                 : Bluetooth Network Connection
        InterfaceDescription : Bluetooth Device (Personal Area Network)
        InterfaceIndex       : 10
        MACAddress           : 5C-F3-70-66-4C-94
        Status               : Disconnected
        DHCPEnabled          : True
        IPAddress            : 169.254.137.77
        DefaultGateway       : 
        SubnetMask           : 255.255.0.0
        DNSServers           : {}
        Messages             : 

        ComputerName         : WORKSTATION-14
        Name                 : eth0
        InterfaceDescription : Intel(R) Ethernet Connection I217-LM
        InterfaceIndex       : 5
        MACAddress           : 34-17-EB-AE-8D-A6
        Status               : Up
        DHCPEnabled          : True
        IPAddress            : 10.60.197.165
        DefaultGateway       : 10.60.196.1
        SubnetMask           : 255.255.254.0
        DNSServers           : {10.15.144.250, 10.8.142.250, 10.11.142.250}
        Messages             : 

    
        
#> 

[CmdletBinding(
    DefaultParameterSetName = "__DefaultParameterSet",
    SupportsShouldProcess = $True,
    ConfirmImpact = "Low"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage="Hostname or FQDN of computer (separate multiple computers with commas)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName = $env:COMPUTERNAME,

    [Parameter(
        HelpMessage = "Return results only for active connections"
    )]
    [Switch] $ConnectedOnly,

    [Parameter(
        HelpMessage = "Return results only for IPv4"
    )]
    [Switch] $IPv4Only,

    [Parameter(
        HelpMessage="Valid credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "Run Invoke-Command scriptblock as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        HelpMessage="Test to run prior to invoke-command (default is WinRM using Kerberos, or Ping, or None)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

    [Parameter(
        Mandatory=$False,
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
    
    # Output
    [array]$Results = @()
    
    #region Scriptblock for invoke-command

    $ScriptBlock = {
        Param($ConnectedOnly,$IPv4Only)
        $ErrorActionPreference = "SilentlyContinue"
        Try {
            $NicOutput = @()
            $OutputTemplate = New-Object PSObject -Property @{
                ComputerName         = $env:COMPUTERNAME
                Name                 = $Null
                InterfaceDescription = $Null
                InterfaceIndex       = $Null
                MACAddress           = $Null
                Status               = $Null
                DHCPEnabled          = $Null
                IPAddress            = $Null
                DefaultGateway       = $Null
                SubnetMask           = $Null
                DNSServers           = $Null
                Messages             = $Null
            }

            $Select = 'ComputerName','Name','InterfaceDescription','InterfaceIndex','MACAddress','Status','DHCPEnabled','IPAddress','DefaultGateway','SubnetMask','DNSServers','Messages'

            # Get the NIC (optionally filter on connected only
            $NICs = Get-NetAdapter -ErrorAction Stop 
            If ($ConnectedOnly.IsPresent) {$NICs = $NICs | Where-Object {$_.MediaConnectionState -eq "Connected"}}
            
            # Foreach NIC
            Foreach ($NICObj in $NICS) {
                
                $Output = $OutputTemplate.PSObject.Copy()
                $Output.Name                 = $NicObj.Name
                $Output.InterfaceDescription = $NICObj.InterfaceDescription
                $Output.InterfaceIndex       = $NicObj.ifIndex
                $Output.MACAddress           = $NICObj.MacAddress
                $Output.Status               = $NICObj.Status
                
                Try {
                    
                    # Get the IP interfaces (optionally filter on IPv4)
                    If ($IPInterfaceObj = $NICObj | Get-NetIPInterface -ErrorAction SilentlyContinue -Verbose:$False) {
                        If ($IPv4Only.IsPresent) {$IPInterfaceObj = $IPInterfaceObj | Where-Object {$_.AddressFamily -eq "IPv4"} }
                        If ($IPInterfaceObj.DHCP -eq "Enabled") {$Output.DHCPEnabled = $True}
                        Else {$Output.DHCPEnabled = $False}
                    }
                     
                   

                    If ($IPObj = ($NICObj | Get-NetIPConfiguration -ErrorAction SilentlyContinue -Verbose:$False)) {
                        Switch ($IPv4Only) {
                            $True {
                                $IPObj = $IPObj | Where-Object {$_.ipv4address}
                                $Output.IPAddress = $IPObj.IPv4Address.IPAddress
                                $Output.DefaultGateway = $IPObj.IPv4DefaultGateway.NextHop
                                $Output.DNSServers = ($IPObj.DNSServer | Where-Object {$_.AddressFamily -eq 2}).ServerAddresses
                            }
                            $False {
                                $Output.IPAddress = $IPObj.IPv4Address.IPAddress,$IPObj.IPv6Address.IPAddress,
                                $Output.DefaultGateway = $IPObj.IPv4DefaultGateway.NextHop,$IPObj.IPv6DefaultGateway.NextHop
                                $Output.DNSServers = $IPObj.DNSServer.ServerAddresses
                            }
                        }
                    }
                        
                    Switch ($IPv4Only) {
                        $True {
                            If ($Index = ($NICObj | Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue -Verbose:$False)) {
                                $Length = $Index.PrefixLength
                                $Output.SubnetMask = (('1'*$Length+'0'*(32-$Length)-split'(.{8})')-ne''| Foreach-Object {[convert]::ToUInt32($_,2)})-join'.'
                            }
                        }
                        $False {
                        
                        }
                    }    

                        
                        
                    }
                }    
                Catch {
                    $Output.Messages = $_.Exception.Message
                }

                $NICOutput += $Output

            } # end foreach
        } # end try getting NICs
        Catch {
            $OutputTemplate.Messages = $_.Exception.Message
            $NICOutput += $OutputTemplate
        }
        
        Write-Output $NICOutput | Select $Select

    } #end scriptblock

    #endregion Scriptblock for invoke-command

    #region Functions

    Function Test-WinRM{
        Param($Computer)
        $Param_WSMAN = @{
            ComputerName   = $Computer
            Credential     = $Credential
            Authentication = "Kerberos"
            ErrorAction    = "Silentlycontinue"
            Verbose        = $False
        }
        Try {
            If (Test-WSMan @Param_WSMAN) {$True}
            Else {$False}
        }
        Catch {$False}
    }

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

    # Splat for Write-Progress
    $Activity = "Get Windows network configuration"
    If ($ConnectedOnly.IsPresent) {$Activity += " (connected adapters only)"}
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as remote PSJob"
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
        ComputerName   = ""
        Authentication = "Kerberos"
        ArgumentList   = $ConnectedOnly
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.AsJob = $True
        $Param_IC.JobName = $Null
        $JobPrefix = "NetIP"
    }
    
    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"Action: $Activity`n")}
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

        If ($Computer -ne $env:COMPUTERNAME) {
            Switch ($ConnectionTest) {
                Default {$Continue = $True}
                Ping {
                    
                    $Msg = "Ping computer"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "Ping failure"
                            $Host.UI.WriteErrorLine("[$Computer] $Msg")
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        $Host.UI.WriteErrorLine("[$Computer] $Msg")
                    }
                }
                WinRM {
                
                    $Msg = "Test WinRM connection"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-WinRM -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "WinRM failure"
                            $Host.UI.WriteErrorLine("[$Computer] $Msg")
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        $Host.UI.WriteErrorLine("[$Computer] $Msg")
                    }
                }
            } #end switch
        } #end if not local computer
        Else {$Continue = $True}

        If ($Continue.IsPresent) {
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Msg = "Invoke command"
                    If ($AsJob.IsPresent) {$Msg += " as PSJob"}
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Jobs += $Job
                    }
                    Else {
                        $Results += Invoke-Command @Param_IC
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                    $Host.UI.WriteErrorLine("[$Computer] $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                $Host.UI.WriteErrorLine("[$Computer] $Msg")
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {

        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            Write-Verbose $Msg
            $Jobs | Get-Job
            
        }
        Else {
            $Msg = "No jobs created"
            Write-Verbose $Msg
        }
    } #end if AsJob

    Else {
        If ($Results.Count -eq 0) {
            $Msg = "No results"
            $Host.UI.WriteWarningLine($Msg)
        }
        Else {
            Write-Output ($Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID)
        }
    }

}

} # end Get-PKWindowsNetConfig

$Null = New-Alias -Name Get-PKNetIPConfig -Value Get-PKWindowsNetConfig -Force -Verbose:$False -Confirm:$False