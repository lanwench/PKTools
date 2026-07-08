#requires -version 4
Function Resolve-PKDNSName {
    <#
.SYNOPSIS
    Performs forward and reverse DNS lookups with optional matching check, on Windows (DNSClient module) or on Mac/Linux (DNSClient-PS module)

.DESCRIPTION
    Performs forward and reverse lookups of one or more names or IP addresses (A, AAAA, CNAME, PTR)
    Optionally tests whether forward and reverse lookups return matching results (-MatchName)
    On Windows, uses the built-in DNSClient module; on Mac/Linux, requires DNSClient-PS from the PS Gallery
    Name defaults to local computer if not specified
    Uses locally-configured nameservers unless -Server is specified
    Accepts pipeline input
    Returns a PSCustomObject


.NOTES
    Name    : Function_Resolve-PKDNSName.ps1
    Created : 2023-03-08
    Author  : Paula Kingsley
    Version : 02.00
    History :

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00 - 2023-03-08 - Created script
        v01.01 - 2023-10-16 - Renamed from Resolve-PKDNS and fixed issue with default name
        v01.02 - 2026-06-16 - Cosmetic updates & removed semantic versioning because OVERKILL yikes
        v02.00 - 2026-06-16 - Added cross-platform support via DnsClient-PS wrapper with normalized output, IP-to-PTR auto-detection, and skip-if-already-imported logic

.PARAMETER Name
    Name or IP address to look up (currently only A/AAAA, CNAME, PTR supported; default is local computername)

.PARAMETER MatchName
    Returns a TRUE/FALSE based on the forward & reverse lookup match

.PARAMETER Detailed
    Returns additional output properties

.PARAMETER Server
    Nameserver (unless specified, uses default locally configured nameserver)

.EXAMPLE
    PS C:\> Resolve-PKDNS -Verbose
        VERBOSE: PSBoundParameters:

        Key              Value
        ---              -----
        Verbose          True
        Name             {LAPTOP14}
        MatchName        False
        Detailed         False
        Server
        ScriptName       Resolve-PKDNSName
        ScriptVersion    02.00
        PipelineInput    False
        ParameterSetName Default


        VERBOSE: [BEGIN: Resolve-PKDNSName] Perform forward/reverse DNS lookups using default nameserver(s)
        VERBOSE: [Prerequisites] Checking for Resolve-DNSName availability
        VERBOSE: [LAPTOP14] Performing lookup
        VERBOSE: [LAPTOP14] 3 result(s) found


        Lookup     : LAPTOP14.domain.local
        Server     : (default)
        RecordType : AAAA
        Name       : LAPTOP14.domain.local
        NameHost   : -
        IPAddress  : fe80::f828:11da:73cb:a954

        Lookup     : LAPTOP14.domain.local
        Server     : (default)
        RecordType : A
        Name       : LAPTOP14.domain.local
        NameHost   : -
        IPAddress  : 10.150.51.143

        Lookup     : LAPTOP14.domain.local
        Server     : (default)
        RecordType : A
        Name       : LAPTOP14.domain.local
        NameHost   : -
        IPAddress  : 172.20.1.185

        VERBOSE: [END: Resolve-PKDNSName] Script ran successfully in 00:00:00.234

.EXAMPLE
    C:\> Get-Content C:\temp\names.txt | Resolve-PKDNS -Verbose

        VERBOSE: PSBoundParameters:

        Key              Value
        ---              -----
        Verbose          True
        Name
        MatchName        False
        Detailed         False
        Server
        ScriptName       Resolve-PKDNSName
        ScriptVersion    02.00
        PipelineInput    True
        ParameterSetName Default

        VERBOSE: [BEGIN: Resolve-PKDNSName] Perform forward/reverse DNS lookups using default nameserver(s)
        VERBOSE: [Prerequisites] Checking for Resolve-DNSName availability
        VERBOSE: [DC14.domain.local] Performing lookup
        VERBOSE: [DC14.domain.local] 1 result(s) found
        Lookup     : DC14.domain.local
        Server     : (default)
        RecordType : A
        Name       : DC14.domain.local
        NameHost   : -
        IPAddress  : 192.168.7.32

        VERBOSE: [192.168.32.8] Performing lookup
        VERBOSE: [192.168.32.8] 1 result(s) found
        Lookup     : 192.168.32.8
        Server     : (default)
        RecordType : PTR
        Name       : 8.32.168.192.in-addr.arpa
        NameHost   : DC11.domain.local
        IPAddress  : -

        VERBOSE: [foo.bar] Performing lookup
        VERBOSE: [foo.bar] 1 result(s) found
        WARNING: [foo.bar] No results returned
        Lookup     : foo.bar
        Server     : (default)
        RecordType : (none)
        Name       : (none)
        NameHost   : (none)
        IPAddress  : (none)

        VERBOSE: [megacorp-net.mail.protection.outlook.com] Performing lookup
        VERBOSE: [megacorp-net.mail.protection.outlook.com] 2 result(s) found
        Lookup     : megacorp-net.mail.protection.outlook.com
        Server     : (default)
        RecordType : A
        Name       : megacorp-net.mail.protection.outlook.com
        NameHost   : -
        IPAddress  : 104.47.71.10

        Lookup     : megacorp-net.mail.protection.outlook.com
        Server     : (default)
        RecordType : A
        Name       : megacorp-net.mail.protection.outlook.com
        NameHost   : -
        IPAddress  : 104.47.71.138

        VERBOSE: [backup-sql.domain.local] Performing lookup
        WARNING: [backup-sql.domain.local] backup-sql.domain.local : DNS name does not exist
        Lookup     : backup-sql.domain.local
        Server     : (default)
        RecordType : (none)
        Name       : (none)
        NameHost   : (none)
        IPAddress  : (none)

        VERBOSE: [END: Resolve-PKDNSName] Script ran successfully in 00:00:01.567

.EXAMPLE
    PS C:\> Resolve-PKDns fe80::f828:11da:73cb:a954,dc9.domain.local -MatchName -Verbose -Detailed

        VERBOSE: PSBoundParameters:

        Key              Value
        ---              -----
        MatchName        True
        Verbose          True
        Detailed         True
        Name             {fe80::f828:11da:73cb:a954, dc9.domain.local}
        Server
        ScriptName       Resolve-PKDNSName
        ScriptVersion    02.00
        PipelineInput    False
        ParameterSetName Match

        VERBOSE: [BEGIN: Resolve-PKDNSName] Perform forward/reverse DNS lookups using default nameserver(s), testing for forward/reverse match
        VERBOSE: [Prerequisites] Checking for Resolve-DNSName availability
        VERBOSE: [fe80::f828:11da:73cb:a954] Performing lookup
        VERBOSE: [fe80::f828:11da:73cb:a954] 1 result(s) found
        VERBOSE: [fe80::f828:11da:73cb:a954] Forward/reverse results match; namehost 'LAPTOP14.domain.local' resolves back to fe80::f828:11da:73cb:a954


        Lookup       : fe80::f828:11da:73cb:a954
        Server       : (default)
        MatchStatus  : True
        RecordType   : PTR
        Section      : Answer
        Name         : 4.2.8.a.b.c.3.7.a.d.1.1.8.2.8.f.0.0.0.0.0.0.0.0.0.0.0.0.0.8.e.f.ip6.arpa.
        TTL          : 1200
        NameHost     : LAPTOP14.domain.local
        IPAddress    : -
        MatchResults : {LAPTOP14.domain.local, LAPTOP14.domain.local, LAPTOP14.domain.local}
        Messages     : {Forward lookup completed successfully, Forward/reverse results match; namehost 'LAPTOP14.domain.local' resolves back to fe80::f828:11da:73cb:a954}

        VERBOSE: [dc9.domain.local] Performing lookup
        VERBOSE: [dc9.domain.local] 1 result(s) found
        VERBOSE: [dc9.domain.local] Testing forward/reverse match for IP address '192.168.30.5'
        VERBOSE: [dc9.domain.local] Forward/reverse lookup results match; IP address '192.168.30.5' resolves back to dc9.domain.local
        Lookup       : dc9.domain.local
        Server       : (default)
        MatchStatus  : True
        RecordType   : A
        Section      : Answer
        Name         : dc9.domain.local
        TTL          : 1200
        NameHost     : -
        IPAddress    : 192.168.30.5
        MatchResults : dc9.domain.local
        Messages     : {Forward lookup completed successfully, Forward/reverse lookup results match; IP address '192.168.30.5' resolves back to dc9.domain.local}

        VERBOSE: [END: Resolve-PKDNSName] Script ran successfully in 00:00:02.345

#>
    [cmdletbinding(DefaultParameterSetName = "Default")]
    Param(
        [Parameter(
            ValueFromPipeline,
            Position = 0,
            HelpMessage = "Name or IP address to look up (currently only A/AAA, CNAME, PTR supported; default is local computername)"
        )]
        [string[]]$Name = $Env:COMPUTERNAME,

        [Parameter(
            ParameterSetName = "Match",
            HelpMessage = "Returns a TRUE/FALSE based on the forward & reverse lookup match"
        )]
        [switch]$MatchName,

        [Parameter(
            HelpMessage = "Returns additional output properties"
        )]
        [switch]$Detailed,

        [Parameter(
            HelpMessage = "Nameserver (unless specified, uses default locally configured nameserver)"
        )]
        [string]$Server
    )
    Begin {

        # Current version (please keep up to date from comment block)
        [version]$Version = "02.00"
        
        $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        # How did we get here?
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $Source = $PSCmdlet.ParameterSetName
        $CurrentParams = $PSBoundParameters
        $ScriptName = $MyInvocation.MyCommand.Name
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path Variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        $CurrentParams.Add("PipelineInput", $PipelineInput)
        $CurrentParams.Add("ParameterSetName", $Source)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
        
        $Activity = "Perform forward/reverse DNS lookups"
        If ($CurrentParams.Server) { $Activity += " via nameserver '$Server'" }
        Else { $Activity += " using default nameserver(s)" }
        If ($MatchName.IsPresent) { $Activity += ", testing for forward/reverse match" }
        Write-Verbose "[BEGIN: $ScriptName] $Activity"

        # Prerequisite check - the DNSClient module is available only on Windows
        Try {
            $Msg = "[Prerequisites] Checking for Resolve-DNSName availability"
            Write-Verbose $Msg
            $Null = Get-Command Resolve-DnsName -ErrorAction Stop
        }
        Catch {
            $Msg = "[Prerequisites] Failed to locate command; looking for DNSClient-PS from the PowerShell Gallery"
            Write-Verbose $Msg
            If (Get-Module -Name DnsClient-PS -ErrorAction SilentlyContinue -Verbose:$False) {
                $Msg = "[Prerequisites] DNSClient-PS already imported; creating wrapper function"
                Write-Verbose $Msg
            }
            ElseIf (Get-Module -Name DnsClient-PS -ListAvailable -ErrorAction SilentlyContinue -Verbose:$False) {
                $Msg = "[Prerequisites] Importing DNSClient-PS module and creating wrapper function"
                Write-Verbose $Msg
                $Null = Import-Module DnsClient-PS -ErrorAction Stop -Verbose:$False
            }
            Else {
                $IsWindows5 = [System.Environment]::OSVersion.Platform -eq 'Win32NT'
                $Msg = If ($IsWindows5 -or $IsWindows) {
                    "Can't locate Resolve-DnsName; ensure the Windows DnsClient module is loaded in the session"
                } Else {
                    "Can't locate Resolve-DnsName!`nInstall the PSGallery cross-platform module using 'Install-Module DnsClient-PS', and then ensure it's loaded in the session"
                }
                Write-Error $Msg
                Break
            }
            Function Resolve-DnsName {
                [CmdletBinding()]
                Param($Name,[string]$Type = 'A',[string]$Server)
                $DnsParam = @{ Query = $Name; QueryType = $Type }
                If ($Server) { $DnsParam.NameServer = $Server }
                # Auto-detect PTR from IP input (mirrors Windows Resolve-DnsName behavior)
                If ($Name -as [System.Net.IPAddress]) {
                    $IP = [System.Net.IPAddress]::Parse($Name)
                    $DnsParam.QueryType = 'PTR'
                    If ($IP.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
                        $Octets = $Name.Split('.')
                        $DnsParam.Query = "$($Octets[3]).$($Octets[2]).$($Octets[1]).$($Octets[0]).in-addr.arpa"
                    }
                    Else {
                        $Bytes = $IP.GetAddressBytes()
                        [Array]::Reverse($Bytes)
                        $Nibbles = $Bytes | ForEach-Object {
                            ('{0:x}' -f ($_ -band 0xF))
                            ('{0:x}' -f (($_ -shr 4) -band 0xF))
                        }
                        $DnsParam.Query = ($Nibbles -join '.') + '.ip6.arpa'
                    }
                } # end check input value type

                $Response = Resolve-Dns @DnsParam

                # Now we'll normalize DnsClient-PS records to match Windows Resolve-DnsName output properties
                #   DnsClient-PS property names assumed: RecordType (enum .ToString() -> "A","AAAA","PTR","CNAME",...),
                #   DomainName (DnsName, implicit string), TimeToLive (int), Address (A/AAAA), PtrDomainName (PTR), CanonicalName (CNAME).
                $SectionMap = [ordered]@{
                    Answers     = 'Answer'
                    Authorities = 'Authority'
                    Additionals = 'Additional'
                }
                Foreach ($SectionKey in $SectionMap.Keys) {
                    Foreach ($Record in $Response.$SectionKey) {
                        $TypeStr = $Record.RecordType.ToString().ToUpper()
                        [PSCustomObject]@{
                            Name      = [string]$Record.DomainName
                            Type      = $TypeStr
                            TTL       = $Record.TimeToLive
                            Section   = $SectionMap[$SectionKey]
                            IPAddress = If ($TypeStr -in @('A', 'AAAA')) { $Record.Address.ToString() } Else { $null }
                            NameHost  = Switch ($TypeStr) {
                                'PTR'   { [string]$Record.PtrDomainName }
                                'CNAME' { [string]$Record.CanonicalName }
                                Default { $null }
                            }
                        }
                    }
                } # end foreach
            } # end wrapper function
        }

        $ResolveParam = @{
            ErrorAction = "Stop"
            Verbose     = $False
        }
        If ($CurrentParams.Server) {$ResolveParam.Add("Server", $Server)}
        $ServerUsed = If ($Server) { $Server } Else { "(default)" }

        If ($MatchName.IsPresent) {
            $Select = "Lookup,Server,MatchStatus,RecordType,Section,Name,TTL,NameHost,IPAddress,MatchResults,Messages" -split (",")
            If (-not $Detailed.IsPresent) { $Select = $Select | Where-Object { $_ -notin @("Section,TTL,MatchResults,Messages" -split (",")) } }
        }
        Else {
            $Select = "Lookup,Server,RecordType,Section,Name,TTL,NameHost,IPAddress,Messages" -split (",")
            If (-not $Detailed.IsPresent) { $Select = $Select | Where-Object { $_ -notin @("Section,TTL,Messages" -split (",")) } }
        }

    }
    Process {
    
        $Total = $Name.Count
        $Current = 0

        Foreach ($Target in $Name) {
        
            $Current ++
            [string]$Msg = "Performing lookup"
            Write-Verbose "[$Target] $Msg"
            If ((-not ($Target -as [ipaddress])) -and ($Target -notmatch "\.")) {
                $Msg = "Hostnames or single-label domains may return errors, but we'll try anyway"
                Write-Warning "[$Target] $Msg"
            } 
            Write-Progress -ID 1 -Activity $Activity -CurrentOperation $Target -PercentComplete ($Current / $Total * 100)
        
            Try {
                [object[]]$Results = Resolve-DNSName -Name $Target @ResolveParam
                $Output = @()
                $Msg = "$($Results.Count) result(s) found"
                Write-Verbose "[$Target] $Msg"

                Foreach ($Result in $Results) {
                    $Messages = @()
                    $MatchResults = $null
                    $MatchStatus = $null
                    $Msg = "Forward lookup completed successfully"
                    $Messages += $Msg

                    Switch ($Result.type) {
                        { $_ -in 'A', 'AAAA' } {
                            If ($MatchName.IsPresent) {
                                $SubActivity = "Testing forward/reverse match for IP address '$($Result.IPAddress)'"
                                Write-Verbose "[$Target] $SubActivity"
                                Write-Progress -id 2 -Activity $SubActivity -CurrentOperation $($Result.IPAddress)
                                $CheckMatch = $null
                                Try {
                                    $CheckMatch = Resolve-PKDNSName -Name $Result.IPAddress @ResolveParam
                                }
                                Catch {
                                    $MatchStatus = "ERROR"
                                    $Msg = "Reverse lookup on resolved IP address failed: $($_.Exception.Message)"
                                    $Messages += $Msg
                                }
                                If ($null -ne $CheckMatch) {
                                    $MatchResults = $CheckMatch.NameHost
                                    If ($MatchResults -contains $Target) {
                                        $MatchStatus = $True
                                        $Msg = "Forward/reverse lookup results match; IP address '$($Result.IPAddress)' resolves back to $($MatchResults -join(", "))"
                                        $Messages += $Msg
                                    }
                                    Else {
                                        $MatchStatus = $False
                                        $Msg = "Forward/reverse lookup results don't match; IP address '$($Result.IPAddress)' resolves back to $($MatchResults -join(", "))"
                                        $Messages += $Msg
                                    }
                                }
                                Write-Verbose "[$Target] $Msg"
                            }
                            $Output += [PSCustomObject]@{
                                Lookup       = $Target
                                Server       = $ServerUsed
                                RecordType   = $Result.Type
                                Section      = $Result.Section
                                TTL          = $Result.TTL
                                Name         = $Result.Name
                                NameHost     = "-"
                                IPAddress    = $Result.IPAddress
                                MatchResults = $MatchResults
                                MatchStatus  = $MatchStatus
                                Messages     = $Messages
                            }
                        }
                        CNAME {
                            If ($MatchName.IsPresent) {
                                $Msg = "No forward/reverse match possible on CNAME records"
                                $Messages += $Msg
                            }
                            $Output += [PSCustomObject]@{
                                Lookup       = $Target
                                Server       = $ServerUsed
                                RecordType   = $Result.Type
                                Section      = $Result.Section
                                TTL          = $Result.TTL
                                Name         = $Result.Name
                                NameHost     = $Result.NameHost
                                IPAddress    = "-"
                                MatchResults = "-"
                                MatchStatus  = "-"
                                Messages     = $Messages
                            }
                        }
                        PTR {
                            If ($MatchName.IsPresent) {
                                $SubActivity = "Testing forward/reverse match for resolved namehost '$($Result.Namehost)'"
                                Write-Progress -id 2 -Activity $SubActivity -CurrentOperation $Target
                                $CheckMatch = $null
                                Try {
                                    $CheckMatch = Resolve-PKDNSName -Name $Result.Namehost @ResolveParam
                                }
                                Catch {
                                    $MatchStatus = "ERROR"
                                    $Msg = "Forward lookup on resolved namehost failed: $($_.Exception.Message)"
                                    $Messages += $Msg
                                }
                                If ($null -ne $CheckMatch) {
                                    $MatchResults = $CheckMatch.Name
                                    If ($CheckMatch | Where-Object { $_.IPAddress -eq $Target }) {
                                        $MatchStatus = $True
                                        $Msg = "Forward/reverse results match; namehost '$($Result.Namehost)' resolves back to $($CheckMatch.IPAddress | Where-Object {$_ -eq $Target})"
                                        $Messages += $Msg
                                    }
                                    Else {
                                        $MatchStatus = $False
                                        $Msg = "Forward/reverse results don't match; namehost '$($Result.Namehost)' resolves back to $($CheckMatch.IPAddress -join(", "))"
                                        $Messages += $Msg
                                    }
                                }
                                Write-Verbose "[$Target] $Msg"
                            }
                            $Output += [PSCustomObject]@{
                                Lookup       = $Target
                                Server       = $ServerUsed
                                RecordType   = $Result.Type
                                Section      = $Result.Section
                                TTL          = $Result.TTL
                                Name         = $Result.Name
                                NameHost     = $Result.NameHost
                                IPAddress    = "-"
                                MatchResults = $MatchResults
                                MatchStatus  = $MatchStatus
                                Messages     = $Messages
                            }
                        }
                        Default {
                            If ($Result.Type -ne 'OPT' -and $MatchName.IsPresent) {
                                $Msg = "Skipping $($Result.Type) record; -MatchName currently supports only A, AAAA, CNAME, and PTR"
                                Write-Warning "[$Target] $Msg"
                            }
                        }
                    } #end switch  
                } #end foreach result

                If ($Output.Count -eq 0) {
                    $Msg = "No results returned"
                    Write-Warning "[$Target] $Msg"
                    $NoResult = "" | Select-Object $Select
                    $NoResult.Lookup = $Target
                    $NoResult.Server = $ServerUsed
                    If ($Detailed.IsPresent) { $NoResult.Messages = $Msg }
                    $NoResult.PSObject.Properties.Name | Foreach-Object { If ($null -eq $NoResult.$_) { $NoResult.$_ = "(none)" } }
                    $Output += $NoResult
                }
                Write-Output $Output | Select-Object $Select
            }
            Catch {
                $Msg = $_.Exception.Message
                Write-Warning "[$Target] $Msg"
                $Output = "" | Select-Object $Select
                $Output.Lookup = $Target
                $Output.Server = $ServerUsed
                If ($Detailed.IsPresent) { $Output.Messages = $Msg }
                $Output.PSObject.Properties.Name | Foreach-Object { If ($null -eq $Output.$_) { $Output.$_ = "(none)" } }
                Write-Output $Output | Select-Object $Select
            }
        } #end foreach name

    }
    End {
        Write-Progress * -Completed
        Write-Verbose "[END: $ScriptName] Script ran successfully in $($Stopwatch.Elapsed.ToString('hh\:mm\:ss\.fff'))"
    }
} #end Resolve-PKDNSName

$Null = New-Alias -Name Resolve-PKDNS -Value Resolve-PKDNSName -Force -ErrorAction SilentlyContinue


