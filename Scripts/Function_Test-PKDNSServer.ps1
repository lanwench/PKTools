#requires -version 3
Function Test-PKDNSServer {
<#
.SYNOPSIS
    Uses Resolve-DNSName to perform record lookups on one or more DNS servers, with optional connectivity tests and option to return lookup output

.DESCRIPTION
    Uses Resolve-DNSName to perform record lookups on one or more DNS servers, with optional connectivity tests and option to return lookup output
    Defaults to locally-configured DNS servers
    Defaults to microsoft.com as lookup target
    By default, attempts to ping DNS server and connect via TCP on port 53, but individual tests can be selected or skipped outright
    Accepts pipeline input for DNS servers
    Returns a PSObject

    Similar functions:
    * Resolve-PKDNSName performs quick name resolution tests using Resolve-DNSName
    * Resolve-PKDNSNameByNET uses the [System.Net.Dns]::GetHostEntryAsync() method for basic lookups on systems without the DNSClient module

.NOTES
    Name    : Function_Test-PKDNSServer.ps1 
    Created : 2019-03-26
    Author  : Paula Kingsley
    Version : 02.00.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK
        
        v01.00.0000 - 2019-03-26 - Created script
        v01.01.0000 - 2019-06-10 - Updated for standardization
        v02.00.0000 - 2019-10-24 - Added option for multiple target lookup names/IPs, changed output properties, 
                                   fixed record/name compatibility test, other updates

.PARAMETER Server
    One or more DNS server names or IPv4 addresses (default is locally configured nameservers)

.PARAMETER Name
    One or more names or IPv4 addresses for DNS lookup test (default is microsoft.com)

.PARAMETER RecordType
    DNS record type: 'Any,' 'A,' 'AAAA,' 'CNAME,' 'MX,' 'NS,' 'PTR,' 'SOA,' 'SRV,' 'TXT' (default is 'A')

.PARAMETER NoRecursion
    Don't perform recursive lookup

.PARAMETER ReturnLookupResults
    Include lookup results in output object (default returns True/False only)

.PARAMETER TruncateLookupOutput
    Truncate lookup results (if returning)

.PARAMETER ConnectionTests
    Perform connection tests prior to lookup: 'Ping,' 'TCP53', 'All,' 'None' (default is 'All')

.PARAMETER ConnectionTestTimeout
    Timeout for ping / TCP connection tests, between 1 millisecond and 30 seconds (default is 5 seconds; ignored if -ConnectionTests is 'None')

.PARAMETER Quiet
    Suppress non-verbose console output


.EXAMPLE
    PS C:\> Test-PKDNSResolution -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                        
        ---                   -----                                        
        Verbose               True                                         
        Server                {10.11.12.13, 10.11.12.17}
        Name                  {microsoft.com}                              
        RecordType            A                                            
        NoRecursion           False                                        
        ReturnLookupResults   False                                        
        TruncateLookupOutput  False                                        
        ConnectionTests       All                                          
        ConnectionTestTimeout 5000                                         
        Quiet                 False                                        
        ScriptName            Test-PKDNSServer                             
        ScriptVersion         2.0.0                                        

        BEGIN  : Test DNS server connectivity and record lookup

        [10.11.12.13] Test ping to DNS server
        [10.11.12.13] Ping succeeded in 4 milliseconds
        [10.11.12.13] Test TCP connection on port 53 to DNS server
        [10.11.12.13] TCP connection on port 53 succeeded in 2 milliseconds
        [10.11.12.13] Perform DNS lookup (record type: A) for 'microsoft.com'
        [10.11.12.13] Host address record lookup succeeded in 40 milliseconds

        Server          : 10.11.12.13
        Name            : microsoft.com
        RecordType      : A
        Recursive       : True
        ConnectionTests : TCP53, Ping
        PingSuccess     : True
        TCP53Success    : True
        LookupSuccess   : True
        LookupResults   : -
        ComputerName    : LAPTOP
        SourceIP        : {192.168.32.6}
        Messages        : Host address record lookup succeeded in 40 milliseconds

        [10.11.12.17] Test ping to DNS server
        [10.11.12.17] Ping succeeded in 10 milliseconds
        [10.11.12.17] Test TCP connection on port 53 to DNS server
        [10.11.12.17] TCP connection on port 53 succeeded in 11 milliseconds
        [10.11.12.17] Perform DNS lookup (record type: A) for 'microsoft.com'
        [10.11.12.17] Host address record lookup succeeded in 17 milliseconds
        
        Server          : 10.11.12.17
        Name            : microsoft.com
        RecordType      : A
        Recursive       : True
        ConnectionTests : TCP53, Ping
        PingSuccess     : True
        TCP53Success    : True
        LookupSuccess   : True
        LookupResults   : -
        ComputerName    : LAPTOP
        SourceIP        : {192.168.32.6}
        Messages        : Host address record lookup succeeded in 17 milliseconds

.EXAMPLE
    PS C:\> Test-PKDNSResolution -Name gmail.com -RecordType A -ReturnLookupResults -Quiet 

        Server          : 10.22.64.7
        Name            : gmail.com
        RecordType      : A
        Recursive       : True
        ConnectionTests : TCP53, Ping
        PingSuccess     : True
        TCP53Success    : True
        LookupSuccess   : True
        LookupResults   : {@{Address=216.58.195.69; IPAddress=216.58.195.69; QueryType=A; IP4Address=216.58.195.69; Name=gmail.com; 
                          Type=A; CharacterSet=Unicode; Section=Answer; DataLength=4; TTL=300}, @{Address=216.58.195.69; 
                          IPAddress=216.58.195.69; QueryType=A; IP4Address=216.58.195.69; Name=gmail.com; Type=A; CharacterSet=Unicode; 
                          Section=Answer; DataLength=4; TTL=300}}
        ComputerName    : WORKSTATION
        SourceIP        : {192.168.5.79}
        Messages        : Host address record lookup succeeded in 42 milliseconds


.EXAMPLE
    PS C:\> Test-PKDNSResolution -Server 4.2.2.1 -Name nyt.com,8.8.8.8 -RecordType A -ConnectionTests None -NoRecursion -ReturnLookupResults -TruncateLookupOutput

        BEGIN  : Test DNS server connectivity and record lookup, and return truncated lookup results

        [4.2.2.1] Perform non-recursive DNS lookup (record type: A) for 'nyt.com'
        [4.2.2.1] Host address record lookup succeeded in 40 milliseconds

        Server          : 4.2.2.1
        Name            : nyt.com
        RecordType      : A
        Recursive       : True
        ConnectionTests : None
        PingSuccess     : -
        TCP53Success    : -
        LookupSuccess   : True
        LookupResults   : @{IP4Address=151.101.193.164; QueryType=A; Name=nyt.com; Type=A}
        ComputerName    : PKTESTWKST
        SourceIP        : {10.28.77.64}
        Messages        : Host address record lookup succeeded in 40 milliseconds

        WARNING: [4.2.2.1] Incompatible lookup target '8.8.8.8' for DNS record type 'A'
        [4.2.2.1] Incompatible record type; operation cancelled by user

        Server          : 4.2.2.1
        Name            : 8.8.8.8
        RecordType      : A
        Recursive       : True
        ConnectionTests : None
        PingSuccess     : -
        TCP53Success    : -
        LookupSuccess   : Error
        LookupResults   : Error
        ComputerName    : PKTESTWKST
        SourceIP        : {10.28.77.64}
        Messages        : Incompatible record type; operation cancelled by user

        END    : Test DNS server connectivity and record lookup, and return truncated lookup results
 
#>
[cmdletbinding(
    SupportsShouldProcess,
    ConfirmImpact = "High"
)]
Param(
    
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more DNS server names or IPv4 addresses (default is locally configured nameservers)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Nameserver","DNSServer","IPAddress","IPv4Address","HostName")]
    [string[]]$Server,

    [Parameter(
        Position = 1,
        HelpMessage = "One or more names or IPv4 addresses for DNS lookup test (default is microsoft.com)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Target","Lookup")]
    [string[]]$Name = "microsoft.com",

    [Parameter(
        HelpMessage = "DNS record type: 'Any,' 'A,' 'AAAA,' 'CNAME,' 'MX,' 'NS,' 'PTR,' 'SOA,' 'SRV,' 'TXT' (default is 'A')"
    )]
    [ValidateSet('Any','A','AAAA','CNAME','MX','NS','PTR','SOA','SRV','TXT')]
    [string]$RecordType = "A",

    [Parameter(
        HelpMessage = "Don't perform recursive lookup"
    )]
    [Switch]$NoRecursion,

    [Parameter(
        ParameterSetName = "Lookup",
        HelpMessage = "Include lookup results in output object (default returns True/False only)"
    )]
    [Alias("ReturnLookupData")]
    [Switch]$ReturnLookupResults,

    [Parameter(
        ParameterSetName = "Lookup",
        HelpMessage = "Truncate lookup results (if returning)"
    )]
    [switch]$TruncateLookupOutput,

    [Parameter(
        HelpMessage = "Perform connection tests prior to lookup: 'Ping,' 'TCP53', 'All,' 'None' (default is 'All')"
    )]
    [ValidateSet("All","Ping","TCP53","None")]
    [String]$ConnectionTests = "All",

    [Parameter(
        HelpMessage = "Timeout for ping / TCP connection tests, between 1 millisecond and 30 seconds (default is 5 seconds; ignored if -ConnectionTests is 'None')"
    )]
    [ValidateRange(1,30000)]
    [int]$ConnectionTestTimeout = 5000,

    [Parameter(
        HelpMessage = "Suppress non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch]$Quiet

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    #$CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    If (-not $CurrentParams.Server) { 
        $Server = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE | Select-Object -Expand DNSServerSearchOrder)
        $CurrentParams.Server = $Server
        
    }
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | Out-String )"

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"

    # Make sure we have the module
    If (-not ($Null = Get-Module DNSClient -ListAvailable -ErrorAction SilentlyContinue -Verbose:$False)) {
        $Msg = "This function requires the DNSClient module, available in Windows 8/2012 and up"
        $Host.UI.WriteErrorLine($Msg)
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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    # Function to check for incompatible lookup/type
    Function Test-CompatibleType {
        [CmdletBinding()]
        Param($RecordType,$Name)
            $ErrorActionPreference = "Stop"
            Try {
                $Msg = "Record type: $RecordType ... Name: $Name"
                Write-Verbose $Msg
                Switch ($RecordType) {
                    PTR {
                        If ([bool]($Name -as [ipaddress])) {$True}
                        Else {$False}
                    }
                    AAAA {
                        # H/T Joakim Svendsen 
                        # As of 2019-03-28: https://www.powershelladmin.com/wiki/PowerShell_.NET_regex_to_validate_IPv6_address_(RFC-compliant)
                        function Test-IsValidIPv6Address {
                            param([Parameter(Mandatory=$true,HelpMessage='Enter IPv6 address to verify')] [string] $IP)
                            $IPv4Regex = '(((25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))'
                            $G = '[a-f\d]{1,4}'
                            # In a case sensitive regex, use: $G = '[A-Fa-f\d]{1,4}'
                            $Tail = @(":",
                                "(:($G)?|$IPv4Regex)",
                                ":($IPv4Regex|$G(:$G)?|)",
                                "(:$IPv4Regex|:$G(:$IPv4Regex|(:$G){0,2})|:)",
                                "((:$G){0,2}(:$IPv4Regex|(:$G){1,2})|:)",
                                "((:$G){0,3}(:$IPv4Regex|(:$G){1,2})|:)",
                                "((:$G){0,4}(:$IPv4Regex|(:$G){1,2})|:)")
                            [string] $IPv6RegexString = $G
                            $Tail | foreach { $IPv6RegexString = "${G}:($IPv6RegexString|$_)" }
                            $IPv6RegexString = ":(:$G){0,5}((:$G){1,2}|:$IPv4Regex)|$IPv6RegexString"
                            $IPv6RegexString = $IPv6RegexString -replace '\(' , '(?:' # make all groups non-capturing
                            [regex] $IPv6Regex = $IPv6RegexString
                            if ($IP -imatch "^$IPv6Regex$") {$true} 
                            else {$false}
                        }
                        If ($Name -match ":") {
                            Test-IsValidIPv6Address -IP $Name
                        }
                        Else {$False}
                    }
                    Any     {$True}
                    Default {
                        If ([bool]($Name -as [ipaddress])) {$False}
                        Else {$True}
                    }
                }
            }
            Catch {
                $False
            }
    } #end Test-CompatibleType

    # Function to ping DNS server
    Function Test-Ping{
        [CmdletBinding()]
        Param($Server,$ConnectionTestTimeout)
        If ($ConnectionTestTimeout) {$Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Server,$ConnectionTestTimeout)}
        Else {$Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Server)}
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
        $Task.Dispose()
    } #end Test-Ping

    # Function to test TCP connection to DNS server
    Function Test-TCP {
        [CmdletBinding()]
        Param($Server,$ConnectionTestTimeout)
        $TCPObject = new-Object system.Net.Sockets.TcpClient
        If ($ConnectionTestTimeout) {
            If ($TCPObject.ConnectAsync($Server,"53").Wait($ConnectionTestTimeout)) {
                $True                
                $TCPObject.Close()
            }
            Else {$False}
        }
        Else {
            If ($TCPObject.ConnectAsync($Server,"53")) {
                $True                
                $TCPObject.Close()
            }
            Else {$False}
        }
    } #end Test-TCP

    # Function to perform DNS lookup
    Function Test-Lookup {
        [CmdletBinding()]
        Param($DNSServer,$RecordType,$Target,[switch]$NoRecurse,[switch]$Boolean,[switch]$Strict,[switch]$Truncate)
        Write-Verbose "[$Server] Look up $RecordType record for '$Target'"
        $Splat = @{
            Name         = $Target
            Server       = $Server
            QuickTimeout = $True
            #DnsOnly      = $True
            TCPOnly      = $True
            NoHostsFile  = $True
            NoRecursion  = $NoRecurse
            ErrorAction  = "SilentlyContinue"
            ErrorVariable = "Fail"
        }
        If ($PSBoundParameters.RecordType) {$Splat.Add("Type",$RecordType)}
        
        $Results = [pscustomobject]@{
            IsResolved = "Error"
            Output     = "Error"
        }

        Try {
            # If we got results
            If ($Resolve = (Resolve-DNSName @Splat| Select *)) {
                
                Write-Verbose ($Resolve | Out-String)    

                # Return fewer properties if truncating (will still get only first row later)
                If ($Truncate.IsPresent) {$Resolve = ($Resolve | Select IP4Address,QueryType,Name,Type)}
                
                # If we want to make sure we are getting the results back only for that record type
                If ($Strict.IsPresent -and $PSBoundParameters.RecordType) {
                    $Resolve = $Resolve | Where-Object {$_.Type -eq $RecordType}
                    If ($Resolve) {
                        $Results.IsResolved = $True
                        $Results.Output = $Resolve
                        Write-Output $Results
                    }
                    Else {
                        $Msg = "Failed to resolve '$Target' with record type '$RecordType'"
                        $Results.IsResolved = $False
                        $Results.Output = $Msg
                    }
                }
                Else {
                    $Results.IsResolved = $True
                    $Results.Output = $Resolve                    
                }
            }
            # If we got nothing...
            Else {
                
                If ($Fail) {
                    $Msg = "FAIL: $($Fail.Exception.Message)"
                }  
                Else {
                    $Msg = "FAIL: Lookup failed"
                }
                $Results.IsResolved = $False
                $Results.Output = $Msg
            }
        }
        Catch {
            $Msg = $_.Exception.Message
            $Results.IsResolved = $False
            $Results.Output = $Msg
        }

        If ($Boolean.IsPresent) {Write-Output $Results.IsResolved}
        Else {Write-Output $Results}

    } #end Test-Lookup

    #endregion Functions
    
    #region Output object

    [switch]$IsRecursive = (-not $NoRecursion)
    $InitialValue = "Error"
    $OutputTemplate = [PSCustomObject]@{
        Server          = $Server
        Name            = $Name
        RecordType      = $RecordType
        Recursive       = $IsRecursive.IsPresent
        ConnectionTests = $InitialValue
        PingSuccess     = $InitialValue
        TCP53Success    = $InitialValue
        LookupSuccess   = $InitialValue
        LookupResults   = $InitialValue
        ComputerName    = $Env:ComputerName
        SourceIP        = @(@(Get-CIMInstance -Class Win32_NetworkAdapterConfiguration -Verbose:$False -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress) -like "*.*") #-join(", ")
        Messages        = $InitialValue
    }
    If (-not $ReturnLookupResults.IsPresent) {
        $OutputTemplate.LookupResults = "-"
    }
    Switch ($ConnectionTests) {
        "None"   {$OutputTemplate.PingSuccess = $OutputTemplate.TCP53Success = "-"}
        "Ping"   {$OutputTemplate.TCP53Success = "-"}
        "TCP53"  {$OutputTemplate.PingSuccess = "-"}
        "All"    {$OutputTemplate.ConnectionTests = "TCP53, Ping"}
    }

    # For the results/output
    $RecordLookup = @{}
    $RecordLookup = @{
        Any   = "DNS record (any type)"
        A     = "Host address record"
        AAAA  = "IPv6 host address record"
        CNAME = "Canonical name record"
        MX    = "Mail eXchange record"
        NS    = "Nameserver record"
        PTR   = "Pointer (reverse) record"
        SOA   = "Start of Authority record"
        SRV   = "Service record"
        TXT   = "Text record"
    }

    #endregion Output object

    #region Splats

    [switch]$TestPing = $True
    [switch]$TestTCPConnection = $True
    Switch ($ConnectionTests) {
        "None" {$TestPing = $TestTCPConnection = $False}
        "Ping" {$TestTCPConnection = $False}
        "TCP"  {$TestPing = $False}
    }

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Test-Ping
    If ($TestPing.IsPresent) {
        $Param_Ping = @{}
        $Param_Ping = @{
            Server                = $Null
            ConnectionTestTimeout = $ConnectionTestTimeout
            ErrorAction           = "Silentlycontinue"
            Verbose               = $False
        }
    }

    # Splat for Test-TCP
    If ($TestTCPConnection.IsPresent) {
        $Param_TCP = @{}
        $Param_TCP = @{
            Server                = $Null
            ConnectionTestTimeout = $ConnectionTestTimeout
            ErrorAction           = "Silentlycontinue"
            Verbose               = $False
        }
    }

    # Splat for Test-Lookup
    $Param_Lookup = @{}
    $Param_Lookup = @{
        DNSServer     = $Null
        Target        = $Name
        Truncate      = $TruncateLookupOutput
        ErrorAction   = "Silentlycontinue"
        Verbose       = $False
        WarningAction = "Silentlycontinue"
        ErrorVariable = "Nope"
    }
    If ($RecordType -ne "Any") {
        $Param_Lookup.Add("RecordType",$RecordType)
        $Param_Lookup.Add("Strict",$True)
    }

    # Splat for Write-Progress
    If ($ConnectionTests -ne "None") {
        $Activity = "Test server connectivity and perform DNS record lookup"
    }
    Else {
        $Activity = "Perform DNS record lookup"
    }
    
    If ($ReturnLookupResults.IsPresent) {
        If ($TruncateLookupOutput.IsPresent) {
            $Activity += ", and return truncated lookup results"
        }
        Else {
            $Activity += ", and return lookup results"
        }
    }
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = $Null
        PercentComplete  = $Null
    }

    #endregion Splats

    # Console output
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}
Process {

    $Total = $Server.Count
    $Current = 0

    Foreach ($S in $Server) {
        
        $Param_WP.Status = $S
        $Param_WP.PercentComplete = ($Current/$Total*100)
        
        $Messages = @()
        [switch]$Continue = $True

        $Output1 = $OutputTemplate.PSObject.Copy()
        $Output1.Server = $S

        # Unless skipping test, ping server; quit on failure
        If ($Continue.IsPresent) {
            
            If ($TestPing.IsPresent) {
            
                #Reset flag
                $Continue = $False

                $Msg = "Test ping to DNS server"
                "[$S] $Msg"  | Write-MessageInfo -FGColor White
                
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP

                Try {
                    $StartTime = Get-Date
                    $Param_Ping.Server = $S
                    If (Test-Ping @Param_Ping) {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "Ping succeeded in $Elapsed milliseconds"
                        "[$S] $Msg"  | Write-MessageInfo -FGColor Green

                        $Output1.PingSuccess = $True

                        # Reset flag
                        $Continue = $True
                    }
                    Else {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "Ping failed after $Elapsed milliseconds"
                        "[$S] $Msg" | Write-MessageError
                        $Output1.PingSuccess = $False
                        Write-Output $Output1
                    }
                }
                Catch {
                    $EndTime = Get-Date
                    $Elapsed = ($EndTime - $StartTime).MilliSeconds
                    $Msg = "Ping failed after $Elapsed milliseconds"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    "[$S] $Msg" | Write-MessageError
                    Write-Output $Output1
                }
            }
        }
        
        # Unless skipping test, connect to server on TCP port 53; quit on failure
        If ($Continue.IsPresent) {

            If ($TestTCPConnection.IsPresent) {

                # Reset flag
                $Continue = $False

                $Msg = "Test TCP connection on port 53 to DNS server"
                "[$S] $Msg"  | Write-MessageInfo -FGColor White
                
                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP

                Try {
                    $StartTime = Get-Date
                    $Param_TCP.Server = $S
                    If ($Null = Test-TCP @Param_TCP) {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "TCP connection on port 53 succeeded in $Elapsed milliseconds"
                        "[$S] $Msg"  | Write-MessageInfo -FGColor Green
                        $Output1.TCP53Success = $True

                        # Reset flag
                        $Continue = $True
                    }
                    Else {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "TCP connection on port 53 failed after $Elapsed milliseconds"
                        "[$S] $Msg" | Write-MessageError
                        $Output1.TCP53Success = $False
                        Write-Output $Output1
                    }
                }
                Catch {
                    $EndTime = Get-Date
                    $Elapsed = ($EndTime - $StartTime).MilliSeconds
                    $Msg = "TCP connection on port 53 failed after $Elapsed milliseconds"
                    If ($ErrorDetails = $_.Exception.Message.Replace('"',"'")) {
                        $Msg += " ($ErrorDetails)"
                    }
                    "[$S] $Msg" | Write-MessageError
                    $Output1.TCP53Success = $False
                    Write-Output $Output1
                }
            }
        }
        
        # Look up record(s)
        If ($Continue.IsPresent) {
            
            $Param_Lookup.DNSServer = $S
                
            Foreach ($N in $Name) {
                
                $N = $N.Trim()
                $Output2 = $Output1.PSObject.Copy()
                $Output2.Name = $N

                # Check to see if we're asking it to resolve itself
                If ($S -eq $N) {
                    
                    # Reset flag
                    $Continue = $False

                    $Msg = "DNS server and lookup target are identical"
                    Write-Warning "[$S] $Msg"
                    
                    $ConfirmMsg = "`n`n`t$Msg!`n`tDo you wish to proceed with DNS query?`n`n"
                    If ($PSCmdlet.ShouldProcess($Server,$ConfirmMsg)) {
                        
                        # Reset flag
                        $Continue = $True
                    }
                    Else {
                        $Msg = "Operation cancelled by user"
                        "[$S] $Msg"  | Write-MessageInfo -FGColor Cyan
                        $Output2.Messages = $Msg
                        Write-Output $Output2
                    }
                }

                If ($Continue.IsPresent) {
                    
                    If ($CurrentParams.RecordType -ne "Any") {

                        Try {
                            If (-not (Test-CompatibleType -RecordType $RecordType -Name $N -ErrorAction Stop)) {

                                # Reset flag
                                $Continue = $False

                                $Msg = "Incompatible lookup target '$N' for DNS record type '$RecordType'"
                                Write-Warning "[$S] $Msg"

                                $ConfirmMsg = "`n`n`t$Msg`n`tDo you wish to proceed with DNS query?`n`n"
                                If ($PSCmdlet.ShouldProcess($Server,$ConfirmMsg)) {
                                
                                    # Reset flag
                                    $Continue = $True
                                }
                                Else {
                                    $Msg = "Operation cancelled by user"
                                    "[$S] $Msg"  | Write-MessageInfo -FGColor Cyan
                                    $Output2.Messages = "Incompatible record type; operation cancelled by user"
                                    Write-Output $Output2
                                }
                            }
                        }
                        Catch {
                            $Msg = "Error testing compatible record type"
                            "[$S] $Msg"  | Write-MessageError
                            $Output2.Messages = $Msg
                            Write-Output $Output2
                        }
                    }
                }

                If ($Continue.IsPresent) {
                
                    If ($NoRecursion.IsPresent) {
                        $Msg = "Perform non-recursive DNS lookup (record type: $RecordType) for '$N'"
                    }
                    Else {
                        $Msg = "Perform DNS lookup (record type: $RecordType) for '$N'"
                    }
                    "[$S] $Msg"  | Write-MessageInfo -FGColor White

                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $Param_Lookup.Target = $N

                    Try {
                        $StartTime = Get-Date
                        $Test = Test-Lookup @Param_Lookup
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                
                        If ($Test.IsResolved -eq $False) {
                            $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                            "[$S] $Msg" | Write-MessageError
                            $Output2.LookupSuccess = $False
                            If ($ReturnLookupResults.IsPresent) {
                                $Output2.LookupResults = $Test.Output
                            }
                            $Output2.Messages = $Msg
                        }
                        Elseif ($Test.Output -match "FAIL: ") {
                            $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                            "[$S] $Msg" | Write-MessageError
                            $Output.LookupSuccess = $False
                            If ($ReturnLookupResults.IsPresent) {
                                $Output2.LookupResults = $Test.Replace("FAIL: ",$Null)
                            }
                            $Output2.Messages = $Msg
                        }
                        Else  {
                            $Msg = "$($RecordLookup[$RecordType]) lookup succeeded in $Elapsed milliseconds"
                            "[$S] $Msg"  | Write-MessageInfo -FGColor Green

                            $Output2.LookupSuccess = $True
                            If ($ReturnLookupResults.IsPresent) {
                                If ($TruncateLookupOutput.IsPresent) {
                                    $Output2.LookupResults = ($Test.Output | Select-Object -First 1) 
                                }
                                Else {
                                    $Output2.LookupResults = $Test.Output
                                }
                            }
                            $Output2.Messages = $Msg   
                        }
                    }
                    Catch {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                        "[$S] $Msg" | Write-MessageError
                        $Output2.Messages = $Msg
                    }

                    Write-Output $Output2
                }
                
            } #end foreach name

        } # end if proceeding to lookup
    }
}
End {
    
    Write-Progress -Activity $Activity -Completed
    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}    

} #end Test-PKDNSServer


$Null = New-Alias Test-PKDNSResolution -Value Test-PKDNSServer -Description "Guessability" -Force -Confirm:$False -ErrorAction SilentlyContinue
