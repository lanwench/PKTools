#requires -Version 3
Function Test-PKNetworkConnections {
<#
.SYNOPSIS
    Performs various connectivity tests to remote computers

.DESCRIPTION
    Performs various connectivity tests to remote computers
    Looks up name in DNS, and defaults to testing: Ping, RDP, Registry, SSH, WinRM and WMI
    Optionally, -Tests parameter brings up menu for selection, including CredSSP
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Test-PKNetworkConnection.ps1
    Created : 2017-11-20
    Author  : Paula Kingsley
    Version : 02.00.0000
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2017-11-20 - Created script based on jrich's original
        v01.00.0001 - 2017-11-30 - Updated help
        v01.01.0000 - 2017-12-21 - DNS test now skipped if IP
        v01.02.0000 - 2018-02-20 - Renamed from Test-PKConnection
        v02.00.0000 - 2019-08-15 - Fixed error with continue flag after DNS, other updates
        
.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Powershell-Test-Server-e0cdea9a?ranMID=24542&ranEAID=je6NUbpObpQ&ranSiteID=je6NUbpObpQ-1p7VW5KEESFnLgSaMed_Bw&tduid=(05cc9a118b47a445311a4d27bb47d63a)(256380)(2459594)(je6NUbpObpQ-1p7VW5KEESFnLgSaMed_Bw)()

.PARAMETER Target
    One or more computer names or IP addresses

.PARAMETER Tests
    Tests to perform: All (DNS, Ping, RDP, Registry, SSH, WinRM, WMI) or ShowMenu (select one or more tests) - default is All

.PARAMETER StopOnError
    Stop testing after first failure

.PARAMETER Credential
    Credentials on target (required for SSH)

.PARAMETER Authentication

.PARAMETER Server
    DNS server(s) (default is locally configured; ignored if DNSClient module is not available)

.PARAMETER Quiet
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Test-PKNetworkConnections -Target ops-pkbastion-1 -Credential $Credential

        BEGIN: Perform 7 network connectivity tests

        [ops-pkbastion-1] Test DNS
        [ops-pkbastion-1] DNS test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test Ping
        [ops-pkbastion-1] Ping test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test RDP
        [ops-pkbastion-1] RDP test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test WinRM
        [ops-pkbastion-1] WinRM test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test Registry
        [ops-pkbastion-1] Registry test succeeded in 1.00 seconds
        [ops-pkbastion-1] Test WMI
        [ops-pkbastion-1] WMI test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test SSH (connection and authentication)
        [ops-pkbastion-1] SSH test failed after 1.00 seconds
        [ops-pkbastion-1] 1 of 7 connection tests failed


        Target         : ops-pkbastion-1
        PercentSuccess : 85.71
        DNS            : 10.62.179.193
        Ping           : True
        WinRM          : True
        Registry       : True
        WMI            : True
        RDP            : True
        SSH            : False
        ElapsedSeconds : 3.00
        Messages       : SSH test failed after 1.00 seconds
                         1 of 7 connection tests failed

        END  : Perform 7 network connectivity tests (Test-PKNetworkConnections completed in 3.00 seconds)


.EXAMPLE

    PS C:\> Test-PKNetworkConnections -Target ops-pkbastion-1 -Tests ShowMenu

        BEGIN: Perform 6 network connectivity tests

        [ops-pkbastion-1] Test DNS
        [ops-pkbastion-1] DNS test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test Ping
        [ops-pkbastion-1] Ping test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test RDP
        [ops-pkbastion-1] RDP test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test WinRM
        [ops-pkbastion-1] WinRM test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test Registry
        [ops-pkbastion-1] Registry test succeeded in 0.00 seconds
        [ops-pkbastion-1] Test WMI
        [ops-pkbastion-1] WMI test succeeded in 0.00 seconds
        [ops-pkbastion-1] 6 of 6 tests completed successfully


        Target         : ops-pkbastion-1
        PercentSuccess : 100
        DNS            : 10.62.179.193
        Ping           : True
        WinRM          : True
        Registry       : True
        WMI            : True
        RDP            : True
        SSH            : -
        ElapsedSeconds : 1.00
        Messages       : 6 of 6 tests completed successfully

        END  : Perform 6 network connectivity tests (Test-PKNetworkConnections completed in 1.00 seconds)

.EXAMPLE
    PS C:\> $arr | Test-PKNetworkConnections -Tests ShowMenu -Verbose | Format-Table -AutoSize
        VERBOSE: PSBoundParameters: 
	
        Key              Value                                    
        ---              -----                                    
        Tests            ShowMenu                                 
        Verbose          True                                     
        Target                                                    
        StopOnError      False                                    
        Credential       System.Management.Automation.PSCredential
        Authentication   Negotiate                                
        Server           {172.19.0.1}                             
        Quiet            False                                    
        SearchFields                                              
        ParameterSetName __AllParameterSets                       
        PipelineInput    True                                     
        ScriptName       Test-PKNetworkConnections                
        ScriptVersion    2.0.0                                    

        VERBOSE: [Prerequisites] 4 test(s) selected: 'DNS', 'Ping', 'RDP', 'WMI'

        BEGIN: Perform 4 network connectivity tests

        [google.com] Test DNS
        [google.com] DNS test succeeded in 0.00 seconds
        [google.com] Test Ping
        [google.com] Ping test succeeded in 0.00 seconds
        [google.com] Test RDP
        [google.com] RDP test failed after 21.00 seconds
        [google.com] Test WMI
        [google.com] WMI test failed after 21.00 seconds
        [google.com] 1 of 4 connection tests failed

        [1.1.1.1] Test DNS
        [1.1.1.1] DNS test succeeded in 0.00 seconds
        [1.1.1.1] Test Ping
        [1.1.1.1] Ping test succeeded in 0.00 seconds
        [1.1.1.1] Test RDP
        [1.1.1.1] RDP test failed after 21.00 seconds
        [1.1.1.1] Test WMI
        [1.1.1.1] WMI test failed after 21.00 seconds
        [1.1.1.1] 1 of 4 connection tests failed
        [192.168.137.1] Test DNS
        [192.168.137.1] DNS test failed after 4.00 seconds
        [192.168.137.1] Test Ping
        [192.168.137.1] Ping test failed after 4.00 seconds
        [192.168.137.1] Test RDP
        [192.168.137.1] RDP test failed after 21.00 seconds
        [192.168.137.1] Test WMI
        [192.168.137.1] WMI test failed after 25.00 seconds
        [192.168.137.1] 1 of 4 connection tests failed

        END  : Perform 4 network connectivity tests (Test-PKNetworkConnections completed in 21.00 seconds)

        Target        PercentSuccess DNS              Ping WinRM Registry   WMI   RDP SSH ElapsedSeconds
        ------        -------------- ---              ---- ----- --------   ---   --- --- --------------
        google.com    75.00          172.217.2.14     True -     -        False False -   42.00         
        1.1.1.1       75.00          one.one.one.one  True -     -        False False -   42.00         
        192.168.137.1 75.00          Error           False -     -        False False -   56.00         

.EXAMPLE
    PS C:\> Test-PKNetworkConnections -Target remote-ws.labdomain.local -Tests ShowMenu -Authentication Kerberos -Server 10.12.8.32 -Quiet

        Target         : remote-ws.labdomain.local
        PercentSuccess : 100
        DNS            : 10.13.68.213
        Ping           : True
        WinRM          : True
        Registry       : -
        WMI            : -
        RDP            : True
        SSH            : -
        ElapsedSeconds : 6.00
        Messages       : 4 of 4 tests completed successfully


#>
[cmdletBinding()]
param(
	[parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more computer names or IP addresses"
    )]
    [Alias("ComputerName","Name","DNSHostName","FQDN","HostName")]
    [ValidateNotNullOrEmpty()]
	[string[]]$Target,

    [parameter(
        HelpMessage = "Tests to perform: All (DNS, Ping, RDP, Registry, SSH, WinRM, WMI) or ShowMenu (select one or more tests) - default is All"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("All","ShowMenu")]
	[string]$Tests = "All",
    
    [Parameter(
        HelpMessage = "Stop testing after first failure"
    )]
    [Switch] $StopOnError,

    [parameter(
        HelpMessage = "Credentials on target (required for SSH)"
    )]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty ,

    [Parameter(
        HelpMessage = "WinRM authentication protocol: Basic, CredSSP, Negotiate, Kerberos, or None (default is Negotiate)"
    )]
    [ValidateSet("None","CredSSP","Basic","Negotiate","Kerberos")]
    [string] $Authentication = "Negotiate",

    [Parameter(
        HelpMessage = "DNS server(s) (default is locally configured)"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]] $Server = (Get-WMIObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'").DNSServerSearchOrder,

    [Parameter(
        HelpMessage = "Hide all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)    
	
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    
    #region Show parameters

    # Display our parameters
    $CurrentParams = $PSBoundParameters
    If ((-not $PipelineInput.IsPresent) -and -not $CurrentParams.Target) {
        $Target = $CurrentParams.Target = $env:COMPUTERNAME
    }

    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.SearchFields = $CurrentParams.SearchFields | Where-Object {$_ -notmatch "AllListed"}
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #endregion Show parameters

    #region Preferences

    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"

    #endregion Preferences

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
    } #end Write-MessageInfo

    # Function to write an error or a warning message
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
        #If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("$Message")}
        #Else {Write-Warning "$Message"}
    } #end Write-MessageError

    # Function to test DNS lookup using Resolve-DNSName, which allows us to specify a nameserver
    Function Test-DNS {
        Param($Name,$Server)
        Try {
            Switch ([bool]($Name -as [ipaddress])) {
                $False {
                    $Prop = "IPAddress"
                    $Type = "A"
                }
                $True {
                    $Prop = "NameHost"
                    $Type = "PTR"
                }
            }
            Resolve-DnsName -Name $Name -Server $Server -Type $Type -Verbose:$False | Select -ExpandProperty $Prop    
        }Catch {}
    } #end Test-DNS

    # Function to test DNS lookup using .NET if DNSClient module isn't available (can't specify a nameserver)
    function Test-DNSNet {
        Param($Name)
        Try {
            Switch ([bool]($Name -as [ipaddress])) {
                $False {If ($Lookup = [Net.Dns]::GetHostAddressesAsync($Name)) {Write-Output $Lookup.Result.IPAddressToString}}
                $True {If ($Lookup = [Net.Dns]::GetHostEntryAsync($Name)) {Write-Output $Lookup.Result.Hostname}}
            }
        }Catch {}
    } #end Test-DNSNet

    # Function to test ping connectivity (faster than test-connection)
    Function Test-Ping{
        Param($Name)
        Try {
            $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Name)
            If ($Task.Result.Status -eq "Success") {$True}
            Else {$False}
            $Task.Dispose()
        }Catch {}
    } #end Test Ping

    # Function to test RDP (TCP 3389)
    Function Test-RDP {
        Param($Computername)
        Try {
            $Socket = New-Object Net.Sockets.TcpClient($ComputerName,3389)
            If ($Socket.Connected -eq $True) {
                $Socket.Close()
                $True
            }
            Else {$False}
        }Catch {$False}
    } #end Test-RDP

    # Function to test registry
    Function Test-Registry {
        Param($ComputerName)
        Try {
            # Keeping old thing for safekeeping, but can't provide credentials, so less useful
            #[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $IP) | Out-Null 
            $Param = @{}
            $Param = @{
                List = $True
                Namespace = "root\default"
            }
            If ($Env:Computername -ne $ComputerName) {
                $Param.Add("ComputerName",$ComputerName)
                $Param.Add("Credential",$Credential)
            }
            #$Reg = Get-WmiObject -List -Namespace root\default -ComputerName $ComputerName -Credential $Credential | Where-Object {$_.Name -eq "StdRegProv"}            
            $Reg = Get-WmiObject @Param | Where-Object {$_.Name -eq "StdRegProv"}
            $HKLM = 2147483650
            If ($Reg.GetStringValue($HKLM,"Software\Microsoft\.NetFramework","InstallRoot").sValue) {$True}
            Else {$False}
        }Catch {$False}
    } #end Test-Registry
    
    # Function to test WinRM connectivity
    Function Test-WinRM{
        Param($ComputerName)
        $Param = @{
            ComputerName   = $ComputerName
            Credential     = $Credential
            Authentication = $Authentication
            ErrorAction    = "Silentlycontinue"
            Verbose        = $False
        }
        Try {
            If (Test-WSMan @Param) {$True}
            Else {$False}
        }
        Catch {$False}
    } #end Test-WinRM

    # Function to test Windows Management Instrumentation
    Function Test-WMI {
        Param($ComputerName)
        Try {
            If (Get-WMIObject -Class Win32_BIOS -ComputerName $ComputerName -Credential $Credential -EA SilentlyContinue) {$True}
            Else {$False}
        }Catch {$False}
    } #end Test-WMI

    # Function to create SSH credential object (strip out any domain stuff from username)
    Function New-SSHCred {
        Switch -regex ($Credential.Username) {
            "\\"    {$Username  = ($Credential.UserName -split("\\"))[1]}
            "@"     {$Username  = ($Credential.UserName -split("@"))[1]}
            Default {$Username  = $Credential.UserName}
        }
        Try {
            $SecureString = $Credential.GetNetworkCredential().Password | ConvertTo-SecureString -asPlainText -Force
            New-Object System.Management.Automation.PSCredential($Username,$SecureString)
        }Catch {}
    }

    # Function to test SSH connection using Posh-SSH module and stored credentials (need to modify to allow keys later)
    Function Test-SSHConnection{
        Param($Computername)
        Try {
            If ($Connect = New-SSHSession -ComputerName $ComputerName -Credential (New-SSHCred) -ConnectionTimeout 50000 -AcceptKey -Force -EA SilentlyContinue -WarningAction SilentlyContinue -Verbose:$False) {
                $Null = $Connect | Remove-SSHSession -EA SilentlyContinue
                $True
            }
            Else {$False}
        }Catch {$False}
    } #end Test-SSHConnection

    # Function to test basic TCP port 22 connectivity (if Posh-SSH not found)
    Function Test-SSHPort {
        Param($ComputerName)
        Try {
            $Connection = New-Object Net.Sockets.TcpClient
            $Connection.Connect($ComputerName,22)
            If ($Connection.Connected) {
                $Connection.Close()
                $Connection = $Null
                $True
            }
            Else {$False}
        }Catch {$False}
    } # end Test-SSHPort

    #endregion Functions
    
    #region Test selection

    [array]$Menu = "DNS","Ping","RDP","Registry","SSH","WinRM","WMI"
    If ($Tests -eq "ShowMenu") {
        If (([array]$Selection = ($Menu | Sort | Out-GridView -Title "Select one or more tests" -OutputMode Multiple)).Count -eq 0){
            $Msg = "At least one selection is mandatory"
            $Msg | Write-MessageError
            Break
        }
        Else {
            $Msg = "$($Selection.Count) test(s) selected: '$($Selection -join("', '"))'"
            Write-Verbose "[Prerequisites] $Msg"
        }
    }
    Else {
        [array]$Selection = $Menu
        $Msg = "$($Selection.Count) test(s) selected: '$($Selection -join("', '"))'"
        Write-Verbose "[Prerequisites] $Msg"
    }

    $TestStr = "test"
    If ($Selection.Count -gt 1) {
        $TestStr = "tests"
    }

    #endregion Test selection

    #region Prerequisites

    # If we don't have Windows 10/2016 we can't use the DNSClient module
    [switch]$LegacyDNS = $False
    If ((([version]((Get-WMIObject Win32_OperatingSystem).Version)).Major -lt 10) -or -not ($Null = Get-Module DnsClient -ErrorAction SilentlyContinue)) {
        $Msg = "Module DNSClient not found (requires Windows version 10 at minimum); DNS lookups will be performed using .NET methods and locally configured DNS servers"
        Write-Warning "[Prerequisites] $Msg"
        $LegacyDNS = $True
    }

    # Make sure we have the necessary files for ssh
    If ($Selection -contains "SSH") {
        [switch]$FullSSH = $True
        If (-not ($Null = Get-Module Posh-SSH)) {
            $Msg = "Posh-SSH module not detected in this session; SSH test will be limited to port 22 connectivity`nSee https://www.powershellgallery.com/packages/Posh-SSH"
            Write-Warning "[Prerequisites] $Msg" 
            $FullSSH = $False
        }
        Else {
            If (-not $Credential.UserName) {
                $Msg = "The Posh-SSH test requires a username and password; please try again with a valid credential"
                "[Prerequisites] $Msg" | Write-MessageError
                Break
            }
        }
    }

    #endregion Prerequisites
    
    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Write-Progress
    $Activity = "Perform $($Selection.Count) network connectivity $TestStr"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }
    
    # Splat for Test-DNS (using Resolve-DNS)
    $Param_DNS = @{}
    $Param_DNS = @{
        Name    = $Null
        Server  = $Server
        Verbose = $False
    }

    # Splat for Test-Registry
    $Param_Reg = @{}
    $Param_Reg = @{
        ComputerName = $Null
        #Credential   = $Credential
    }

    # Splat for Test-WinRM
    $Param_WinRM = @{}
    $Param_WinRM = @{
        ComputerName   = $Null
        Authentication = $Authentication
        Credential     = $Credential
    }

    # Splat for Test-SSHConnection
    $Param_SSH = @{}
    $Param_SSH = @{
        Computername = $Null
    }

    # Splat for Test-WMI
    $Param_WMI = @{}
    $Param_WMI = @{
        ComputerName   = $Null
        Credential     = $Credential
    }

    #endregion Splats
        
    #region  Output

    $InitialValue = "Error"
    $OutputTemplate = @()
    $OutputTemplate = [pscustomobject]@{
        Target         = $InitialValue
        PercentSuccess = $InitialValue
        DNS            = $InitialValue
        Ping           = $InitialValue
        WinRM          = $InitialValue
        Registry       = $InitialValue
        WMI            = $InitialValue
        RDP            = $InitialValue
        SSH            = $InitialValue
        ElapsedSeconds = $InitialValue
        Messages       = $InitialValue
    }

    #endregion Output
    
    # Timer for all tests
    $Stopwatch_Script =  [system.diagnostics.stopwatch]::StartNew()

    # Console output
    $Msg = "BEGIN: $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title
}

Process{
    
    $Total = $Target.Count
    $Current = 0

    Foreach ($T in $Target){
        
        $Current ++
        $Param_WP.Status = $T
        $Param_WP.PercentComplete = ($Current / $Total * 100)

        $Output = $OutputTemplate.PSObject.Copy()
        $Output.Target = $T
        $Failures = 0
        $Messages = @()

        # Timer for all tests
        $Stopwatch_Tests =  [system.diagnostics.stopwatch]::StartNew()

        [switch]$Continue = $True

        #region DNS lookup

        $Label = "DNS"
        If ($Selection -contains $Label -and $Continue.IsPresent) {
        
            # Reset flag
            If ($StopOnError.IsPresent) {$Continue = $False}
        
	        Try {
	        
                $Msg = "Test $Label"
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
                "[$T] $Msg" | Write-MessageInfo -FGColor White

                Switch ($LegacyDNS.IsPresent) {
                    $True {
                        ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                        $DNSResult = Test-DNSNet -Name $T 
                        ${"$Stopwatch_$Label"}.Stop()
                    }
                    $False {       
                        $Param_DNS.Name = $T
                        ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                        $DNSResult = Test-DNS @Param_DNS
                        ${"$Stopwatch_$Label"}.Stop()
                    }
                }
                If ($DNSResult) {
                    $Output.DNS = $DNSResult
                    $Msg = "$Label test succeeded in $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageInfo -FGColor Green

                    $Continue = $True
                }
                Else {
                    $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageError
                    $Messages += $Msg
                    $Output.DNS = "(Lookup failed)"
                    $Failures ++
                }
	        }
	        Catch {
		        $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$T] $Msg" | Write-MessageError
                $Messages += $Msg
	            $Failures ++
	        }
        } #end DNS
        Else {
            $Output.DNS = "-"
        }

        #endregion DNS lookup

        #region Test ping

        $Label = "Ping"
        If ($Selection -contains $Label -and $Continue.IsPresent) {
            
            # Reset flag
            If ($StopOnError.IsPresent) {$Continue = $False}

            Try {
                
                $Msg = "Test $Label"
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
                "[$T] $Msg" | Write-MessageInfo -FGColor White

                ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                $Ping = (Test-Ping -Name $T)
                ${"$Stopwatch_$Label"}.Stop()

                If ($Ping) {
                    $Output.Ping = $True
                    $Msg = "$Label test succeeded in $("{0:N2}" -f ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageInfo -FGColor Green
                        
                    $Continue = $True
                }
                Else  {
                    $Output.Ping = $False
                    $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageError
                    $Messages += $Msg
                    $Failures ++
                }
            }
            Catch {
                $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$T] $Msg" | Write-MessageError
                $Messages += $Msg
                $Failures ++
            }
        } # end ping
        Else {
            $Output.Ping = "-"
        }

        #endregion Test ping

        #region Test Remote Desktop Protocol

        $Label = "RDP"
	    If ($Selection -contains $Label -and $Continue.IsPresent) {
    	    
            # Reset flag
            If ($StopOnError.IsPresent) {$Continue = $False}

		    Try {
                $Msg = "Test $Label"
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
                "[$T] $Msg" | Write-MessageInfo -FGColor White

                ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                $RDP = Test-RDP -Computername $T
                ${"$Stopwatch_$Label"}.Stop()
                
                If ($RDP) {
                    $Output.RDP = $True
                    
                    $Msg = "$Label test succeeded in $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageInfo -FGColor Green

                    $Continue = $True
                }
		        Else {
			        $Output.RDP = $False
                    $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageError
                    $Messages += $Msg
                    $Failures ++
		        }
            }
            Catch {
                $Msg = "$Label test failed after $("{0:N2}" -f ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$T] $Msg" | Write-MessageError
                $Messages += $Msg
                $Failures ++
                
            }
        } #end RDP
        Else {
            $Output.RDP = "-"
        }

        #endregion Test Remote Desktop Protocol

        #region Test Windows Remote Management

        $Label = "WinRM"
	    If ($Selection -contains $Label -and $Continue.IsPresent) {
    	    
            # Reset flag
            If ($StopOnError.IsPresent) {$Continue = $False}

		    Try {
                
                $Msg = "Test $Label"
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
                "[$T] $Msg" | Write-MessageInfo -FGColor White

                $Param_WinRM.ComputerName = $T

                ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                $WinRM = Test-WinRM @Param_WinRM
                ${"$Stopwatch_$Label"}.Stop()
                
                If ($WinRM) {
                    $Output.WinRM = $True
                    $Msg = "$Label test succeeded in $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageInfo -FGColor Green
                }
		        Else {
			        $Output.WinRM = $False
                    $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageError
                    $Messages += $Msg
                    $Failures ++
		        }
            }
            Catch {
                $Msg = "$Label test failed after $("{0:N2}" -f ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$T] $Msg" | Write-MessageError
                $Messages += $Msg
                $Failures ++
                
            }
        } #end WinRM
        Else {
            $Output.WinRM = "-"
        }
        
        #endregion Test Windows Remote Management

        #region Test remote registry lookup

        $Label = "Registry"   
        If ($Selection -contains $Label -and $Continue.IsPresent) {    
            
             # Reset flag
            If ($StopOnError.IsPresent) {$Continue = $False}

            Try {
                
                $Msg = "Test $Label"
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
                "[$T] $Msg" | Write-MessageInfo -FGColor White

                $Param_Reg.ComputerName = $T

                ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                $Registry = Test-Registry @Param_Reg
                ${"$Stopwatch_$Label"}.Stop()

                If ($Registry) {
                    $Output.Registry = $True
                    $Msg = "$Label test succeeded in $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageInfo -FGColor Green
                }
		        Else {
			        $Output.Registry = $False
                    $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageError
                    $Messages += $Msg
                    $Failures ++
		        }
			}
			Catch {
                $Msg = "$Label test failed on $T"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$T] $Msg" | Write-MessageError
                $Messages += $Msg
                $Failures ++
            }
        }
        Else {
            $Output.Registry = "-"
        }

        #endregion Test remote registry lookup

        # RPC https://blogs.technet.microsoft.com/runcmd/troubleshoot-rpc-with-powershell/ 

        #region Test WMI connection

        $Label = "WMI"
        If ($Selection -contains $Label -and $Continue.IsPresent) {    
            
            # Reset flag
            If ($StopOnError.IsPresent) {$Continue = $False}
            
            Try {	
                $Msg = "Test $Label"
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
                "[$T] $Msg" | Write-MessageInfo -FGColor White
            
                $Param_WMI.ComputerName = $T

                ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                $WMI = Test-WMI @Param_WMI
                ${"$Stopwatch_$Label"}.Stop()

                If ($WMI) {
                    $Output.WMI = $True
                    $Msg = "$Label test succeeded in $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageInfo -FGColor Green

                    $Continue = $True
                }
		        Else {
			        $Output.WMI = $False
                    $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageError
                    $Messages += $Msg
                    $Failures ++
		        }
			}
			Catch {
                $Msg = "$Label test failed on $T"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$T] $Msg" | Write-MessageError
                $Messages += $Msg
                $Failures ++
            }
        }
        Else {
            $Output.WMI = "-"
        }
        
        #endregion Test WMI connection        
            
        #region Test SSH connection

        $Label = "SSH"
        If ($Selection -contains $Label -and $Continue.IsPresent) {    
            
            # Reset flag
            If ($StopOnError.IsPresent) {$Continue = $False}
            
            Try {	
                $Msg = "Test $Label"
                Switch ($FullSSH.IsPresent) {
                    $True  {$Msg += " (connection and authentication)"}
                    $False {$Msg += " (TCP connection only)"}
                }
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
                "[$T] $Msg" | Write-MessageInfo -FGColor White
                
                Switch ($FullSSH) {
                    $True  {
                        $Param_SSH.ComputerName = $T
                        ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                        $SSH = Test-SSHConnection @Param_SSH
                        ${"$Stopwatch_$Label"}.Stop()
                    }
                    $False {
                        ${"$Stopwatch_$Label"} =  [system.diagnostics.stopwatch]::StartNew()
                        $SSH = Test-SSHPort -ComputerName $T
                        ${"$Stopwatch_$Label"}.Stop()
                    }
                }

                If ($SSH) {
                    $Output.SSH = $True
                    $Msg = "$Label test succeeded in $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageInfo -FGColor Green

                    $Continue = $True
                }
		        Else {
			        $Output.SSH = $False
                    $Msg = "$Label test failed after $("{0:N2}" -f  ${"$Stopwatch_$Label"}.Elapsed.Seconds) seconds"
                    "[$T] $Msg" | Write-MessageError
                    $Messages += $Msg
                    $Failures ++
		        }
			}
			Catch {
                $Msg = "$Label test failed on $T"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$T] $Msg" | Write-MessageError
                $Messages += $Msg
                $Failures ++
            }
        }
        Else {
            $Output.SSH = "-"
        }

        #endregion Test SSH connection

        # Total time
        $Stopwatch_Tests.Stop()
        $Completed = $("{0:N2}" -f  ($Stopwatch_Tests.Elapsed.Seconds))
        $Msg = "Testing completed in $Completed seconds"
        $Output.ElapsedSeconds = $Completed


        If ($Failures -gt 0) {
            $PercentSuccess = $("{0:N2}" -f (100 - ($Failures.Count/$Selection.Count * 100)))
            $Output.PercentSuccess = $PercentSuccess
            $Msg = "$($Failures.Count) of $($Selection.Count) connection $Teststr failed"
            $Messages += $Msg
            $Color = "Yellow"
        }
        Else {
            $Output.PercentSuccess = 100
            $Msg = "$($Selection.Count) of $($Selection.Count) $TestStr completed successfully"
            $Color = "Green"
            $Messages += $Msg
        }

        "[$T] $Msg" | Write-MessageInfo -FGColor $Color
        $Output.Messages = $Messages -join("`n")
        Write-Output $Output

    } #end foreach target

}
End{
    
    $Stopwatch_Script.Stop()
    $Msg = "$($MyInvocation.MyCommand.Name) completed in $("{0:N2}" -f  $Stopwatch_Script.Elapsed.Seconds) seconds"

    $Msg = "END  : $Activity ($Msg)"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

    Write-Progress -Activity $Activity -Completed

}
} #end Test-PKNetworkConnection