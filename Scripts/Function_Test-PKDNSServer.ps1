#requires -Version 3
Function Test-PKDNSServer {
<#
.SYNOPSIS
    Uses Resolve-DNSName to test lookups on a DNS server, with connection tests and option to return lookup output

.DESCRIPTION
    Uses Resolve-DNSName to test lookups on a DNS server, with connection tests and option to return lookup output
    Defaults to locally-configured DNS servers
    Defaults to microsoft.com as lookup target
    By default attempts to ping DNS server and connect via TCP on port 53, but individual tests can be selected or skipped outright
    Accepts pipeline input
    Returns a PSObject

.NOTES
    Name    : Function_Test-PKDNSServer.ps1 
    Created : 2019-03-26
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK
        
        v01.00.0000 - 2019-03-26 - Created script

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
    Suppress non-verbose/non-error console output

.EXAMPLE
    
#>
[cmdletbinding(
    SupportsShouldProcess,
    ConfirmImpact = "Medium"
)]
Param(
    [Parameter(
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "DNS server name or IP (default is local)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Nameserver","DNSServer","IPAddress","IPv4Address")]
    [string[]]$Server,
    
    [Parameter(
        HelpMessage = "Name or IP for DNS lookup (default is microsoft.com)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Target","Name")]
    [string]$Lookup = "microsoft.com",

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
        HelpMessage = "Include lookup data in output object"
    )]
    [Switch]$ReturnLookupData,

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
    [version]$Version = "01.00.0000"
    
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
        Exit
    }

    #region Functions

    # Function to check for incompatible lookup/type
    Function Test-CompatibleType {
        [CmdletBinding()]
        Param($RecordType,$Lookup)
            $ErrorActionPreference = "Stop"
            Try {
                Switch ($RecordType) {
                    PTR {
                        If ($Lookup -as [ipaddress]) {$True}
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
                        If ($Lookup -match ":") {
                            Test-IsValidIPv6Address -IP $Lookup
                        }
                        Else {$False}
                    }
                    Any     {$True}
                    Default {
                        If ($Lookup -as [ipaddress]) {$False}
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

    # Function to test telnet to DNS server
    Function Test-Telnet {
        [CmdletBinding()]
        Param($Server,$ConnectionTestTimeout)
        $TCPObject = new-Object system.Net.Sockets.TcpClient
        If ($ConnectionTestTimeout) {
            If ($TCPObject.ConnectAsync($Server,"53").Wait($ConnectionTestTimeout)) {
                $True                
                $TCPObject.Close()
            }
            Else {
                $False
            }
        }
        Else {
            If ($TCPObject.ConnectAsync($Server,"53")) {
                $True                
                $TCPObject.Close()
            }
            Else {
                $False
            }
        }
    } #end Test-Telnet

    # Function to test A record lookup on DNS server
    Function Test-Lookup {
        [CmdletBinding()]
        Param($Server,$RecordType,$Lookup,[switch]$NoRecurse,[switch]$Boolean,[switch]$Strict)
        Write-Verbose "[$Server] Look up $RecordType record for '$Lookup'"
        $Param = @{
            Name         = $Lookup
            Server       = $Server
            QuickTimeout = $True
            DNSOnly      = $True
            TCPOnly      = $True
            NoHostsFile  = $True
            NoRecursion  = $NoRecurse
            ErrorAction  = "SilentlyContinue"
        }
        If ($PSBoundParameters.RecordType) {$Param.Add("Type",$RecordType)}
        Try {
            # If we got results
            If ($Results = (Resolve-DNSName @Param | Select *)) {
                Write-Verbose ($Results | Out-String)

                # If we want to make sure we are getting the results back only for that record type
                If ($Strict.IsPresent -and $PSBoundParameters.RecordType) {
                    If ($Results.QueryType -eq $RecordType) {
                        If ($Boolean.IsPresent) {$True}
                        Else {Write-Output $Results}
                    }
                    Else {
                        If ($Boolean.IsPresent) {$False}
                        Else {Write-Output "Failed to resolve '$Lookup'"}
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
                Else {Write-Output "Lookup failed"}   
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
            If (-not (Test-CompatibleType -RecordType $RecordType -Lookup $Lookup -ErrorAction Stop)) {
                $Msg = "Incompatible lookup target '$Lookup' for DNS record type '$RecordType'"
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
    $OutputTemplate = New-Object PSObject -Property ([ordered]@{
        Server         = $Null
        Lookup         = $Lookup
        RecordType     = $RecordType
        Recursive      = $IsRecursive
        TestPing       = $InitialValue
        TestConnection = $InitialValue
        TestLookup     = $InitialValue
        Output         = $InitialValue
        ComputerName   = $Env:ComputerName
        SourceIP       = @(@(Get-WmiObject Win32_NetworkAdapterConfiguration | Select-Object -ExpandProperty IPAddress) -like "*.*")        
        Messages       = $InitialValue

    })
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
        Any   = "DNS record (any available)"
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

    # Splat for Test-Ping
    If ($TestTCPConnection.IsPresent) {
        $Param_Telnet = @{}
        $Param_Telnet = @{
            Server                = $Null
            ConnectionTestTimeout = $ConnectionTestTimeout
            ErrorAction           = "Silentlycontinue"
            Verbose               = $False
        }
    }

    # Splat for Test-Lookup
    $Param_Lookup = @{}
    $Param_Lookup = @{
        Server        = $Null
        Lookup        = $Lookup
        Boolean       = $True
        ErrorAction   = "Silentlycontinue"
        Verbose       = $False
        WarningAction = "Silentlycontinue"
        ErrorVariable = "Nope"
    }
    If ($ReturnLookupData.IsPresent) {$Param_Lookup.Boolean = $False}
    If ($RecordType -ne "Any") {
        $Param_Lookup.Add("RecordType",$RecordType)
        $Param_Lookup.Add("Strict",$True)
    }

    # Splat for Write-Progress
    $Activity = "Test DNS lookup"
    If ($ReturnLookupData.IsPresent) {$Activity += " and return lookup results"}
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = $Null
        PercentComplete  = $Null
    }

    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg`n")}
    Else {Write-Verbose $Msg}

}
Process {

    $Total = $DNSServer.Count
    $Current = 0

    Foreach ($S in $DNSServer) {
        
        $Param_WP.Status = $S
        $Param_WP.PercentComplete = ($Current/$Total*100)
        
        $Messages = @()
        [switch]$Continue = $False

        $Output = $OutputTemplate.PSObject.Copy()
        $Output.Server = $S

        # Just in case someone hasn't been careful with their parameters (ask me how I know)
        If ($Server -eq $Lookup) {
            $Msg = "DNS server and lookup target are identical"
            $Messages += $Msg
            Write-Warning "[$S] $Msg"
            $ConfirmMsg = "`n$Msg`nDo you wish to proceed with DNS query?`n"
            If ($PSCmdlet.ShouldContinue($ConfirmMsg,$Server)) {
                $Continue = $True
            }
            Else {
                $Msg = "Operation cancelled by user"
                $Messages += $Msg
                If (-not $Quiet.IsPresent) {
                    $FGColor = "White"
                    $Host.UI.WriteLine($FGColor,$BGColor,"[$S] $Msg")
                }
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
                If (-not $Quiet.IsPresent) {
                    $FGColor = "White"
                    $Host.UI.WriteLine($FGColor,$BGColor,"[$S] $Msg")
                }
                Else {Write-Verbose "[$S] $Msg"}

                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP

                Try {
                    $StartTime = Get-Date
                    $Param_Ping.Server = $S
                    If (Test-Ping @Param_Ping) {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "Ping succeeded in $Elapsed milliseconds"
                        If (-not $Quiet.IsPresent) {
                            $FGColor = "Green"
                            $Host.UI.WriteLine($FGColor,$BGColor,"[$S] $Msg")
                        }
                        $Output.TestPing = $True
                        $Messages += $Msg
                        $Continue = $True
                    }
                    Else {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Output.TestPing = $False
                        $Msg = "Ping failed after $Elapsed milliseconds"
                        $Host.UI.WriteErrorLine("[$S] $Msg")
                        $Messages += $Msg
                    }
                }
                Catch {
                    $EndTime = Get-Date
                    $Elapsed = ($EndTime - $StartTime).MilliSeconds
                    $Msg = "Ping failed after $Elapsed milliseconds"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine("[$S] $Msg")
                    $Messages += $Msg
                }
            }
            Else {
                #$Output.TestPing = "-"
                $Msg = "Ping test skipped"
                Write-Verbose "[$S] $Msg"
                $Messages += $Msg
            }
        }
        
        # Test TCP connection
        If ($Continue.IsPresent) {

            If ($TestTCPConnection.IsPresent) {

                # Reset flag
                $Continue = $False

                $Msg = "Test TCP connection on port 53 to DNS server"
                If (-not $Quiet.IsPresent) {
                    $FGColor = "White"
                    $Host.UI.WriteLine($FGColor,$BGColor,"[$S] $Msg")
                }
                Else {Write-Verbose "[$S] $Msg"}

                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP

                Try {
                    $StartTime = Get-Date
                    $Param_Telnet.Server = $S
                    If ($Null = Test-Telnet @Param_Telnet) {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "TCP connection on port 53 succeeded in $Elapsed milliseconds"
                        If (-not $Quiet.IsPresent) {
                            $FGColor = "Green"
                            $Host.UI.WriteLine($FGColor,$BGColor,"[$S] $Msg")
                        }
                        $Output.TestConnection = $True
                        $Messages += $Msg
                        $Continue = $True
                    }
                    Else {
                        $EndTime = Get-Date
                        $Elapsed = ($EndTime - $StartTime).MilliSeconds
                        $Msg = "TCP connection on port 53 failed after $Elapsed milliseconds"
                        $Host.UI.WriteErrorLine("[$S] $Msg")
                        $Output.TestConnection = $False
                        $Messages += $Msg
                    }
                }
                Catch {
                    $EndTime = Get-Date
                    $Elapsed = ($EndTime - $StartTime).MilliSeconds
                    $Msg = "TCP connection on port 53 failed after $Elapsed milliseconds"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine("[$S] $Msg")
                    $Messages += $Msg
                }
            }
            Else {
                #$Output.TestConnection = "-"
                $Msg = "Connection test skipped"
                Write-Verbose "[$S] $Msg"
                $Messages += $Msg
            }
        }
        

        # Look up A record
        If ($Continue.IsPresent) {
            
            $Msg = "Test record lookup for '$Lookup' on DNS server"
            If ($NoRecursion.IsPresent) {$Msg += " (non-recursive)"}
            If (-not $Quiet.IsPresent) {
                $FGColor = "White"
                $Host.UI.WriteLine($FGColor,$BGColor,"[$S] $Msg")
            }
            Else {Write-Verbose "[$S] $Msg"}

            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP

            $Param_Lookup.Server = $S
            Try {
                $StartTime = Get-Date
                $Test = Test-Lookup @Param_Lookup
                $EndTime = Get-Date
                $Elapsed = ($EndTime - $StartTime).MilliSeconds
                
                If ($Test -in ($False,"Lookup failed")) {
                    $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                    $Host.UI.WriteErrorLine("[$S] $Msg")
                    $Output.TestLookup = $False
                    $Messages += $Msg
                }
                Else {
                    $Msg = "$($RecordLookup[$RecordType]) lookup succeeded in $Elapsed milliseconds"
                    $Output.TestLookup = $True
                    $FGColor = "Green"
                    If (-not $Quiet.IsPresent) {
                        $FGColor = "Green"
                        $Host.UI.WriteLine($FGColor,$BGColor,"[$S] $Msg")
                    }
                    If ($ReturnLookupData.IsPresent) {
                        $Output.Output = $Test | Out-String
                    }
                    $Messages += $Msg
                }
            }
            Catch {
                $EndTime = Get-Date
                $Elapsed = ($EndTime - $StartTime).MilliSeconds
                $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("[$S] $Msg")
                $Messages += $Msg
            }
        } # end if proceeding to lookup
        
        $Output.Messages = $Messages -join ("`n")
        Write-Output $Output
    }

}
End {
    
    Write-Progress -Activity $Activity -Completed

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg")}
    Else {Write-Verbose $Msg}

}    

} #end Test-PKDNSServer


$Null = New-Alias Test-PKDNSResolution -Value Test-PKDNSServer -Description "Guessability" -Force -Confirm:$False -ErrorAction SilentlyContinue