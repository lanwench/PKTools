#requires -Version 3
Function Test-PKDNSServer {
<#
.SYNOPSIS
    Uses Resolve-DNSName to perform DNS lookups on one or more servers, with connectivity tests and option to return lookup output

.DESCRIPTION
    Uses Resolve-DNSName to perform DNS lookups on one or more servers, with connectivity tests and option to return lookup output
    Defaults to locally-configured DNS servers
    Defaults to microsoft.com as lookup target
    By default attempts to ping DNS server and connect via TCP on port 53, but individual tests can be selected or skipped outright
    Accepts pipeline input for DNS servers
    Returns a PSObject

.NOTES
    Name    : Function_Test-PKDNSServer.ps1 
    Created : 2019-03-26
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK
        
        v01.00.0000 - 2019-03-26 - Created script
        v01.01.0000 - 2019-06-10 - Updated for standardization

.PARAMETER Server
    DNS server name or IP (default is locally configured DNS servers)

.PARAMETER Lookup
    Name or IP address for record lookup (default is microsoft.com)

.PARAMETER RecordType
    DNS record type: Any, A, AAAA, CNAME, MX, NS, PTR, SOA, SRV, TXT (default is 'A')

.PARAMETER NoRecursion
    Don't perform recursive lookup

.PARAMETER ReturnLookupData
    Include lookup data in output object

.PARAMETER ConnectionTests
    Perform connection tests prior to lookup: Ping, TCP port 53, All, None (default is All)

.PARAMETER ConnectionTestTimeout
    Timeout for ping / TCP connection tests, between 1 millisecond and 30 seconds (default is 5 seconds; ignored if -ConnectionTests is 'None')

.PARAMETER Quiet
    Suppress non-verbose console output

.EXAMPLE
    PS C:\> Test-PKDNSServer -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value           
        ---                   -----           
        Server                {172.30.41.8} 
        Verbose               True            
        Name                  microsoft.com   
        RecordType            A               
        NoRecursion           False           
        ReturnLookupData      False           
        TruncateLookupOutput  False           
        ConnectionTests       All             
        ConnectionTestTimeout 5000            
        Quiet                 False           
        ScriptName            Test-PKDNSServer
        ScriptVersion         1.1.0           

        VERBOSE: This function performs a DNS lookup for a single name against multiple DNS servers using the Resolve-DNS
        Name cmdlet (part of the DNSClient module) with the option to test connectivity and return Boolean output for the
         query.
        The function Resolve-PKDNSName resolves multiple names against multiple DNS servers, with no connectivity tests.
        The function Resolve-PKDNSNameByNET uses the [System.Net.Dns]::GetHostEntryAsync() method for basic lookups on sy
        stems without the DNSClient module.

        BEGIN  : Test DNS server connectivity and record lookup

        [172.30.41.8] Test ping to DNS server
        [172.30.41.8] Ping succeeded in 10 milliseconds
        [172.30.41.8] Test TCP connection on port 53 to DNS server
        [172.30.41.8] TCP connection on port 53 succeeded in 10 milliseconds
        [172.30.41.8] Perform DNS lookup (record type: A) for 'microsoft.com'
        [172.30.41.8] Host address record lookup succeeded in 29 milliseconds

        Server         : 172.30.41.8
        Name           : microsoft.com
        IsResolved     : True
        RecordType     : A
        Recursive      : True
        TestPing       : True
        TestConnection : True
        ComputerName   : LAPTOP
        SourceIP       : {172.30.55.79}
        Messages       : Host address record lookup succeeded in 29 milliseconds

        END    : Test DNS server connectivity and record lookup

.EXAMPLE
    PS C:\> Test-PKDNSServer -Server 1.1.1.1 -Name cloudflare.com -RecordType NS -NoRecursion -ReturnLookupData -ConnectionTests TCP -Quiet

        Server         : 1.1.1.1
        Name           : cloudflare.com
        IsResolved     : True
        RecordType     : NS
        Recursive      : False
        TestPing       : -
        TestConnection : True
        Output         : {@{QueryType=NS; Server=ns3.cloudflare.com; NameHost=ns3.cloudflare.com; Name=cloudflare.com; 
                         Type=NS; CharacterSet=Unicode; Section=Answer; DataLength=46; TTL=21459}, @{QueryType=NS; 
                         Server=ns4.cloudflare.com; NameHost=ns4.cloudflare.com; Name=cloudflare.com; Type=NS; 
                         CharacterSet=Unicode; Section=Answer; DataLength=46; TTL=21459}, @{QueryType=NS; 
                         Server=ns5.cloudflare.com; NameHost=ns5.cloudflare.com; Name=cloudflare.com; Type=NS; 
                         CharacterSet=Unicode; Section=Answer; DataLength=46; TTL=21459}, @{QueryType=NS; 
                         Server=ns6.cloudflare.com; NameHost=ns6.cloudflare.com; Name=cloudflare.com; Type=NS; 
                         CharacterSet=Unicode; Section=Answer; DataLength=46; TTL=21459}...}
        ComputerName   : LAPTOP
        SourceIP       : {172.30.55.79}
        Messages       : Nameserver record lookup succeeded in 86 milliseconds

.EXAMPLE
    PS C :\> Test-PKDNSServer -Server 99.99.99.99 -Name cloudflare.com -RecordType NS -NoRecursion -ReturnLookupData -ConnectionTests TCP -Quiet

        Server         : 99.99.99.99
        Name           : cloudflare.com
        IsResolved     : Error
        RecordType     : NS
        Recursive      : False
        TestPing       : -
        TestConnection : False
        Output         : Error
        ComputerName   : LAPTOP
        SourceIP       : {172.30.55.79}
        Messages       : TCP connection on port 53 failed after 2 milliseconds

    
#>
[cmdletbinding(
    SupportsShouldProcess,
    ConfirmImpact = "Medium"
)]
Param(
    
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "DNS server name or IP (default is locally configured)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Nameserver","DNSServer","IPAddress","IPv4Address","HostName")]
    [string[]]$Server,

    [Parameter(
        Position = 1,
        HelpMessage = "Name or IP for DNS lookup test (default is microsoft.com)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Target","Lookup")]
    [string]$Name = "microsoft.com",

    [Parameter(
        HelpMessage = "DNS record type: Any, A, AAAA, CNAME, MX, NS, PTR, SOA, SRV, TXT (default is 'A')"
    )]
    [ValidateSet('Any','A','AAAA','CNAME','MX','NS','PTR','SOA','SRV','TXT')]
    [string]$RecordType = "A",

    [Parameter(
        HelpMessage = "Don't perform recursive lookup"
    )]
    [Switch]$NoRecursion,

    [Parameter(
        ParameterSetName = "Lookup",
        HelpMessage = "Include lookup data in output object (default is True/False only)"
    )]
    [Switch]$ReturnLookupData,

    [Parameter(
        ParameterSetName = "Lookup",
        HelpMessage = "Truncate lookup data if returning"
    )]
    [switch]$TruncateLookupOutput,

    [Parameter(
        HelpMessage = "Perform connection tests prior to lookup: Ping, TCP port 53, All, None (default is All)"
    )]
    [ValidateSet("All","Ping","TCP","None")]
    [String]$ConnectionTests = "All",

    [Parameter(
        HelpMessage = "Timeout for Ping / TCP connection tests, between 1 millisecond and 30 seconds (default is 5 seconds; ignored if -ConnectionTests is 'None')"
    )]
    [ValidateRange(1,30000)]
    [int]$ConnectionTestTimeout = 5000,

    [Parameter(
        HelpMessage = "Suppress non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch]$Quiet

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"
    
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

    # Function options
    $Msg = "This function performs a DNS lookup for a single name against multiple DNS servers using the Resolve-DNSName cmdlet (part of the DNSClient module) with the option to test connectivity and return Boolean output for the query.`nThe function Resolve-PKDNSName resolves multiple names against multiple DNS servers, with no connectivity tests.`nThe function Resolve-PKDNSNameByNET uses the [System.Net.Dns]::GetHostEntryAsync() method for basic lookups on systems without the DNSClient module.`n"
    Write-Verbose $Msg

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

    # Function to check for incompatible lookup/type
    Function Test-CompatibleType {
        [CmdletBinding()]
        Param($RecordType,$Name)
            $ErrorActionPreference = "Stop"
            Try {
                Switch ($RecordType) {
                    PTR {
                        If ($Name -as [ipaddress]) {$True}
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
                        If ($Name -as [ipaddress]) {$False}
                        Else {$True}
                    }
                }
            }
            Catch {
                break
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
        Try {
            # If we got results
            If ($Results = (Resolve-DNSName @Splat| Select *)) {
                Write-Verbose ($Results | Out-String)

                # Return fewer properties if truncating (will still get only first row later)
                If ($Truncate.IsPresent) {$Results = ($Results | Select IP4Address,QueryType,Name,Type)}
                
                # If we want to make sure we are getting the results back only for that record type
                If ($Strict.IsPresent -and $PSBoundParameters.RecordType) {
                    $Results = $Results | Where-Object {$_.Type -eq $RecordType}
                    If ($Results) {
                        If ($Boolean.IsPresent) {$True}
                        Else {Write-Output $Results}
                    }
                    Else {
                        If ($Boolean.IsPresent) {$False}
                        Else {Write-Output "Failed to resolve '$Target' with record type '$RecordType'"}
                    }
                }
                Else {
                    If ($Boolean.IsPresent) {$True}
                    Else {Write-Output $Results}
                }
            }
            # If we got nothing...
            Else {
                If ($Boolean.IsPresent) {$False}
                Else {
                    If ($Fail) {Write-Output "FAIL: $($Fail.Exception.Message)"}   
                    Else {Write-Output "FAIL: Lookup failed"}
                }
            }
        }
        Catch {
            $Msg = $_.Exception.Message
            If ($Boolean.IsPresent) {$False}
            Else {Write-Output $Msg}
        }
    } #end Test-Lookup

    #endregion Functions
    
    #region Test compatibility of lookup name & record type

    If ($CurrentParams.RecordType -ne "Any") {
        Try {
            If (-not (Test-CompatibleType -RecordType $RecordType -Lookup $Name -ErrorAction Stop)) {
                $Msg = "Incompatible lookup target '$Name' for DNS record type '$RecordType'"
                $Host.UI.WriteErrorLine("$Msg")
                Break
            }
        }
        Catch {}
    }

    #endregion Test compatibility of lookup name & record type

    #region Output object

    [switch]$IsRecursive = (-not $NoRecursion)
    $InitialValue = "Error"
    $OutputTemplate = [PSCustomObject]@{
        Server         = $Server
        Name           = $Name
        IsResolved     = $InitialValue
        RecordType     = $RecordType
        Recursive      = $IsRecursive.IsPresent
        TestPing       = $InitialValue
        TestConnection = $InitialValue
        #TestLookup     = $InitialValue
        Output         = $InitialValue
        ComputerName   = $Env:ComputerName
        SourceIP       = @(@(Get-CIMInstance -Class Win32_NetworkAdapterConfiguration -Verbose:$False -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress) -like "*.*") #-join(", ")
        Messages       = $InitialValue
    }
    If (-not $ReturnLookupData.IsPresent) {
        $OutputTemplate.PSObject.Properties.Remove("Output")
    }
    Switch ($ConnectionTests) {
        "None" {$OutputTemplate.TestPing = $OutputTemplate.TestConnection = "-"}
        "Ping" {$OutputTemplate.TestConnection = "-"}
        "TCP"  {$OutputTemplate.TestPing = "-"}
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
        $Activity = "Test DNS server connectivity and record lookup"
    }
    Else {
        $Activity = "Test DNS server lookup"
    }
    
    If ($ReturnLookupData.IsPresent) {
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
        [switch]$Continue = $False

        $Output = $OutputTemplate.PSObject.Copy()
        $Output.Server = $S

        # Just in case someone hasn't been careful with their parameters (ask me how I know)
        If ($Server -eq $Name) {
            $Msg = "DNS server and lookup target are identical"
            "[$S] $Msg" | Write-MessageInfo -FGColor Red
            $ConfirmMsg = "`n$Msg`nDo you wish to proceed with DNS query?`n"
            If ($PSCmdlet.ShouldContinue($ConfirmMsg,$Server)) {
                $Continue = $True
            }
            Else {
                $Msg = "Operation cancelled by user"
                "[$S] $Msg"  | Write-MessageInfo -FGColor Cyan
            }
        }
        Else {
            $Continue = $True
        }

        # Test ping
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

                        $Output.TestPing = $True
                        $Continue = $True
                    }
                    Else {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "Ping failed after $Elapsed milliseconds"
                        "[$S] $Msg" | Write-MessageError

                        $Output.TestPing = $False
                    }
                }
                Catch {
                    $EndTime = Get-Date
                    $Elapsed = ($EndTime - $StartTime).MilliSeconds
                    $Msg = "Ping failed after $Elapsed milliseconds"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    "[$S] $Msg" | Write-MessageError
                }
            }
            Else {
                $Msg = "Ping test skipped"
                "[$S] $Msg"  | Write-MessageInfo -FGColor Cyan
            }
        }
        
        # Test TCP connection
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
                        $Output.TestConnection = $True
                        $Continue = $True
                    }
                    Else {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "TCP connection on port 53 failed after $Elapsed milliseconds"
                        "[$S] $Msg" | Write-MessageError
                        $Output.TestConnection = $False
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
                    $Output.TestConnection = $False
                }
            }
            Else {
                $Msg = "Connection test skipped"
                "[$S] $Msg"  | Write-MessageInfo -FGColor Cyan
            }
        }
        

        # Look up A record
        If ($Continue.IsPresent) {
            
            If ($NoRecursion.IsPresent) {
                $Msg = "Perform non-recursive DNS lookup (record type: $RecordType) for '$Name'"
            }
            Else {
                $Msg = "Perform DNS lookup (record type: $RecordType) for '$Name'"
            }
            "[$S] $Msg"  | Write-MessageInfo -FGColor White

            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP

            $Param_Lookup.DNSServer = $S
            Try {
                $StartTime = Get-Date
                $Test = Test-Lookup @Param_Lookup
                $EndTime = Get-Date
                $Elapsed = ($EndTime - $StartTime).MilliSeconds
                
                If ($Test -eq $False) {
                    $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                    "[$S] $Msg" | Write-MessageError
                    $Output.IsResolved = $False
                    $Output.Messages = $Msg
                }
                Elseif ($Test -match "FAIL: ") {
                    $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                    "[$S] $Msg" | Write-MessageError
                    $Output.IsResolved = $False
                    If ($ReturnLookupData.IsPresent) {
                        $Output.Output = $Test.Replace("FAIL: ",$Null)
                    }
                    $Output.Messages = $Msg
                }
                Else  {
                    $Msg = "$($RecordLookup[$RecordType]) lookup succeeded in $Elapsed milliseconds"
                    "[$S] $Msg"  | Write-MessageInfo -FGColor Green

                    $Output.IsResolved = $True
                    If ($ReturnLookupData.IsPresent) {
                        If ($TruncateLookupOutput.IsPresent) {
                            $Output.Output = ($Test | Select-Object -First 1) 
                        }
                        Else {
                            $Output.Output = $Test
                        }
                        }
                    $Output.Messages = $Msg   
                }

                <#

                If ($Test -in ($False,"Lookup failed")) {
                    $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                    "[$S] $Msg" | Write-MessageError
                    $Output.TestLookup = $False
                }
                Else {
                    $Msg = "$($RecordLookup[$RecordType]) lookup succeeded in $Elapsed milliseconds"
                    $Output.TestLookup = $True
                    "[$S] $Msg"  | Write-MessageInfo -FGColor Green

                    If ($ReturnLookupData.IsPresent) {
                        If ($TruncateLookupOutput.IsPresent) {
                            $Output.Output = ($Test | Select-Object -First 1) # -join("`n"))# | Out-String)
                        }
                        Else {
                            $Output.Output = $Test # | Out-String
                        }
                    }
                }

                #>
            }
            Catch {
                $EndTime = Get-Date
                $Elapsed = ($EndTime - $StartTime).MilliSeconds
                $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                "[$S] $Msg" | Write-MessageError
            }
        } # end if proceeding to lookup
        
        $Output.Messages = $Msg
        Write-Output $Output
    }

}
End {
    
    Write-Progress -Activity $Activity -Completed
    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}    

} #end Test-PKDNSServer


$Null = New-Alias Test-PKDNSResolution -Value Test-PKDNSServer -Description "Guessability" -Force -Confirm:$False -ErrorAction SilentlyContinue
