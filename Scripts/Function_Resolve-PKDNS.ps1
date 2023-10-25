#requires -version 4
Function Resolve-PKDNSName {
    <#
.SYNOPSIS 
    Performs forward and reverse lookups of one or more names or IP addresses, optionally testing for forward/reverse name match

.DESCRIPTION
    Performs forward and reverse lookups of one or more names or IP addresses, optionally testing for forward/reverse name match
    Defaults to local computer if -Name is not specified
    Uses locally-configured nameservers unles -Server is specified
    Accepts pipeline input
    Returns a PSObject

.NOTES
    Name    : Function_Resolve-PKDNSName.ps1
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2023-03-08 - Created script
        v01.01.0000 - 2023-10-16 - Renamed from Resolve-PKDNS and fixed issue with default name

.PARAMETER Name
    Name or IP address to look up (currently only A/AAA, CNAME, PTR supported; default is local computername)

.PARAMETER MatchName
    Returns a TRUE/FALSE based on the forward & reverse lookup match

.PARAMETER Detailed
    Returns additional output properties

.PARAMETER Server
    Nameserver (unless specified, uses default locally configured nameserver)

.EXAMPLE
    PS C:\ Resolve-PKDNS -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key              Value        
        ---              -----        
        Verbose          True         
        Name             {LAPTOP14}
        MatchName        False        
        Detailed         False        
        Server                        
        ScriptName       Resolve-PKDNS
        ScriptVersion    1.0.0        
        PipelineInput    False        
        ParameterSetName Default      


        VERBOSE: [BEGIN: Resolve-PKDNS] Perform forward/reverse DNS lookups against default nameserver(s)
        VERBOSE: [LAPTOP14] Performing lookup
        VERBOSE: [LAPTOP14] 3 result(s) found


        Lookup     : LAPTOP14.domain.local
        RecordType : AAAA
        Name       : LAPTOP14.domain.local
        NameHost   : -
        IPAddress  : fe80::f828:11da:73cb:a954

        Lookup     : LAPTOP14.domain.local
        RecordType : A
        Name       : LAPTOP14.domain.local
        NameHost   : -
        IPAddress  : 10.150.51.143

        Lookup     : LAPTOP14.domain.local
        RecordType : A
        Name       : LAPTOP14.domain.local
        NameHost   : -
        IPAddress  : 172.20.1.185

        VERBOSE: [END: Resolve-PKDNS]

.EXAMPLE
    C:\> get-content C:\temp\names.txt | Resolve-PKDNS -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value        
        ---              -----        
        Verbose          True         
        Name                          
        MatchName        False        
        Detailed         False        
        Server                        
        ScriptName       Resolve-PKDNS
        ScriptVersion    1.0.0        
        PipelineInput    True         
        ParameterSetName Default      

        VERBOSE: [BEGIN: Resolve-PKDNS] Perform forward/reverse DNS lookups against default nameserver(s)
        VERBOSE: [DC14.domain.local] Performing lookup
        VERBOSE: [DC14.domain.local] 1 result(s) found
        Lookup     : DC14.domain.local
        RecordType : A
        Name       : DC14.domain.local
        NameHost   : -
        IPAddress  : 192.168.7.32

        VERBOSE: [192.168.32.8] Performing lookup
        VERBOSE: [192.168.32.8] 1 result(s) found
        Lookup     : 192.168.32.8
        RecordType : PTR
        Name       : 8.32.168.192.in-addr.arpa
        NameHost   : DC11.domain.local
        IPAddress  : -

        VERBOSE: [foo.bar] Performing lookup
        VERBOSE: [foo.bar] 1 result(s) found
        VERBOSE: [foo.bar] Processing SOA record foo.bar
        Lookup     : foo.bar
        RecordType : SOA
        Name       : ns0.centralnic-dns.com
        NameHost   : -
        IPAddress  : -

        VERBOSE: [megacorp-net.mail.protection.outlook.com] Performing lookup
        VERBOSE: [megacorp-net.mail.protection.outlook.com] 2 result(s) found
        Lookup     : megacorp-net.mail.protection.outlook.com
        RecordType : A
        Name       : megacorp-net.mail.protection.outlook.com
        NameHost   : -
        IPAddress  : 104.47.71.10

        Lookup     : megacorp-net.mail.protection.outlook.com
        RecordType : A
        Name       : megacorp-net.mail.protection.outlook.com
        NameHost   : -
        IPAddress  : 104.47.71.138

        VERBOSE: [backup-sql.domain.local] Performing lookup
        WARNING: [backup-sql.domain.local] backup-sql.domain.local : DNS name does not exist
        Lookup     : backup-sql.domain.local
        RecordType : ERROR
        Name       : ERROR
        NameHost   : ERROR
        IPAddress  : ERROR

        VERBOSE: [END: Resolve-PKDNS]

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
        ScriptName       Resolve-PKDNS                                      
        ScriptVersion    1.0.0                                              
        PipelineInput    False                                              
        ParameterSetName Match                                              

        VERBOSE: [BEGIN: Resolve-PKDNS] Perform forward/reverse DNS lookups against default nameserver(s), testing for forward/reverse match
        VERBOSE: [fe80::f828:11da:73cb:a954] Performing lookup
        VERBOSE: [fe80::f828:11da:73cb:a954] 1 result(s) found
        VERBOSE: [fe80::f828:11da:73cb:a954] Testing forward/reverse match for resolved namehost 'LAPTOP14.domain.local'
        VERBOSE: [fe80::f828:11da:73cb:a954] Forward/reverse match; forward lookup on resolved namehost 'LAPTOP14.domain.local' equals or includes: fe80::f828:11da:73cb:a954


        Lookup       : fe80::f828:11da:73cb:a954
        RecordType   : PTR
        Section      : Question
        Name         : 4.2.8.a.b.c.3.7.a.d.1.1.8.2.8.f.0.0.0.0.0.0.0.0.0.0.0.0.0.8.e.f.ip6.arpa.
        TTL          : 1200
        NameHost     : LAPTOP14.domain.local
        IPAddress    : -
        MatchResults : {fe80::f828:11da:73cb:a954, 10.11.51.30, 172.20.1.185}
        IsMatch      : True
        Messages     : Forward/reverse match; forward lookup on resolved namehost 'LAPTOP14.domain.local' equals or includes: fe80::f828:11da:73cb:a954

        VERBOSE: [dc9.domain.local] Performing lookup
        VERBOSE: [dc9.domain.local] 1 result(s) found
        VERBOSE: [dc9.domain.local] Testing forward/reverse match for IP address '192.168.30.5'
        VERBOSE: [dc9.domain.local] Forward/reverse match; reverse lookup on resolved IP address '192.168.30.5' returns: dc9.domain.local
        Lookup       : dc9.domain.local
        RecordType   : A
        Section      : Answer
        Name         : dc9.domain.local
        TTL          : 1200
        NameHost     : -
        IPAddress    : 192.168.30.5
        MatchResults : dc9.domain.local
        IsMatch      : True
        Messages     : Forward/reverse match; reverse lookup on resolved IP address '192.168.30.5' returns: dc9.domain.local

        VERBOSE: [END: Resolve-PKDNS]

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
        [version]$Version = "01.01.0000"

        # How did we get here?
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $Source = $PSCmdlet.ParameterSetName
        $CurrentParams = $PSBoundParameters

        <#
        If (-not $CurrentParams.Name -and -not $PipelineInput.IsPresent) {
            $Msg = "Getting local computer name or FQDN"
            Write-Verbose "[Prerequisites] $Msg"
            # Let's not be lazy; try to get the FQDN of the local computer so we aren't relying on search domains
            $Name = (Get-WmiObject -Class Win32_ComputerSystem -PipelineVariable CS -ErrorAction SilentlyContinue | 
                Select-Object @{N = "Name"; E = {
                        If ($CS.Domain) { "$($CS.DNSHostName).$($CS.Domain)" }
                        Else { $Env:ComputerName }
                    }
                }).Name
            $CurrentParams.Name = $Name
        }
        #>

        #If ($PipelineInput.IsPresent) {$CurrentParams.Name = $Null}

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
    
        $Param = @{
            ErrorAction = "Stop"
            Verbose     = $False
        }
        If ($CurrentParams.Server) {$Param.Add("Server", $Server)}

        If ($MatchName.IsPresent) {
            $Select = "Lookup,RecordType,Section,Name,TTL,NameHost,IPAddress,MatchResults,IsMatch,Messages" -split (",")
            If (-not $Detailed.IsPresent) { $Select = $Select | Where-Object { $_ -notin @("Section,TTL,MatchResults,Messages" -split (",")) } }
        }
        Else {
            $Select = "Lookup,RecordType,Section,Name,TTL,NameHost,IPAddress,Messages" -split (",")
            If (-not $Detailed.IsPresent) { $Select = $Select | Where-Object { $_ -notin @("Section,TTL,Messages" -split (",")) } }
        }

        # Not currently in use
        #Function PTRtoIP($PTR) {
        #    # Thanks, Boe Prox
        #    $T = ($PTR -replace '.in-addr.arpa$').split(".")
        #    $T[-1..-($T.Count)] -join '.'
        #}

    
        $Activity = "Perform forward/reverse DNS lookups"
        If ($CurrentParams.Server) { $Activity += " against nameserver '$Server'" }
        Else { $Activity += " against default nameserver(s)" }
        If ($MatchName.IsPresent) { $Activity += ", testing for forward/reverse match" }
        Write-Verbose "[BEGIN: $ScriptName] $Activity"

    }
    Process {
    
        $Total = $Name.Count
        $Current = 0

        Foreach ($Target in $Name) {
        
            $Current ++
            If ((-not $Target -as [ipaddress]) -and (-not $Target -match "\.")) {
                $Msg = "Hostnames or single-label domains may return errors"
                Write-Warning "[$Target] $Msg"
            } 
            [string]$Msg = "Performing lookup"
            # Write-Verbose "[$Target] $Msg"
            Write-Progress -ID 1 -Activity $Activity -CurrentOperation $Target -PercentComplete ($Current / $Total * 100)
        
            Try {
                [object[]]$Results = Resolve-DNSName -Name $Target @Param
                $Output = @()
                $Msg = "$($Results.Count) result(s) found"
                Write-Verbose "[$Target] $Msg"
                Foreach ($Result in $Results) {
                    
                    $Messages = @()
                    $Msg = "Forward lookup completed successfully"
                    $Messages += $Msg

                    Switch ($Result.type) {
                        { $_ -in 'A', 'AAAA' } {
                            
                            #$Msg = "Forward lookup completed successfully"
                            #$Messages += $Msg

                            If ($MatchName.IsPresent) {
                                $Activity = "Testing forward/reverse match for IP address" # $($Result.IPAddress)"
                                #Write-Verbose "[$Target] $Msg"
                                Write-Progress -id 2 -Activity $Activity -CurrentOperation $($Result.IPAddress) # -PercentComplete ($Current / $Total * 100)
                                If ($CheckMatch = Resolve-PKDNSName -Name $Result.IPAddress @Param) {
                                    $MatchResults = $CheckMatch.NameHost
                                    If ($MatchResults -contains $Target) {
                                        $IsMatch = $True
                                        $Msg = "Forward/reverse lookup results match; IP address '$($Result.IPAddress)' resolves to $($MatchResults -join(", "))"
                                        $Messages += $Msg
                                    }
                                    Else {
                                        $IsMatch = $False
                                        $Msg = "Forward/reverse lookup results don't match; IP address '$($Result.IPAddress)' returns: $($MatchResults -join(", "))"
                                        $Messages += $Msg
                                    }
                                }
                                Else {
                                    $IsMatch = "ERROR"
                                    $Msg = "Reverse lookup on resolved IP address failed"
                                    $Messages += $Msg
                                }
                                Write-Verbose "[$Target] $Msg"
                            }
                            $Output += [PSCustomObject]@{
                                Lookup       = $Target
                                RecordType   = $Result.Type
                                Section      = $Result.Section
                                TTL          = $Result.TTL
                                Name         = $Result.Name
                                NameHost     = "-"
                                IPAddress    = $Result.IPAddress
                                MatchResults = $MatchResults            
                                IsMatch      = $IsMatch
                                Messages     = $Messages -join("`n")
                            }
                        }
                        CNAME {
                            If ($MatchName.IsPresent) {
                                $Msg = "No forward/reverse match possible on CNAME records"
                                $Messages += $Msg
                            }
                            $Output += [PSCustomObject]@{
                                Lookup       = $Target
                                RecordType   = $Result.Type
                                TTL          = $Result.TTL
                                Name         = $Result.Name
                                NameHost     = $Result.NameHost
                                IPAddress    = "-"
                                MatchResults = "-"
                                IsMatch      = "-"
                                Messages     = $Messages -join("`n")
                            }
                        }
                        PTR {
                            If ($MatchName.IsPresent) {
                                $Activity = "Testing forward/reverse match for resolved namehost '$($Result.Namehost)'"
                                Write-Progress -id 2 -Activity $Activity -CurrentOperation $Target
                                If ($CheckMatch = Resolve-PKDNSName -Name $Result.Namehost @Param) {
                                    $MatchResults = $CheckMatch.Name
                                    If ($CheckMatch | Where-Object { $_.IPAddress -eq $Target }) {
                                        $IsMatch = $True
                                        $Msg = "Forward/reverse results match; namehost '$($Result.Namehost)' resolves to $($CheckMatch.IPAddress | Where-Object {$_ -eq $Target})"
                                        $Messages += $Msg
                                    }
                                    Else {
                                        $IsMatch = $False
                                        $Msg = "Forward/reverse results don't match; namehost '$($Result.Namehost)' resolves to $($CheckMatch.IPAddress -join(", "))"
                                        $Messages += $Msg
                                    }
                                }
                                Else {
                                    $IsMatch = "ERROR"
                                    $Msg = "Forward lookup on resolved namehost failed"
                                    $Messages += $Msg
                                    
                                }
                                Write-Verbose "[$Target] $Msg"
                            }
                            $Output += [PSCustomObject]@{
                                Lookup       = $Target
                                RecordType   = $Result.Type
                                TTL          = $Result.TTL
                                Section      = $Result.Section
                                Name         = $Result.Name
                                NameHost     = $Result.NameHost
                                IPAddress    = "-"
                                MatchResults = $MatchResults
                                IsMatch      = $IsMatch 
                                Messages     = $Messages -join("`n")
                            }
                        }
                        Default {
                            If ($MatchName.IsPresent) { 
                                $Msg = "Skipping $($Result.Type) record; -MatchName currently supports only A, AA, CNAME, and PTR" 
                                Write-Warning "[$Target] $Msg"
                            }
                        }
                    } #end switch  
                } #end foreach result

                Write-Output $Output | Select-Object $Select
            }
            Catch {
                $Msg = $_.Exception.Message
                Write-Warning "[$Target] $Msg"
                $Output = "" | Select-Object $Select
                $Output.Lookup = $Target
                If ($Detailed.IsPresent) { $Output.Messages = $Msg }
                $Output.PSObject.Properties.Name | Foreach-Object { If ($Null = $Output.$_) { $Output.$_ = "ERROR" } }
                Write-Output $Output | Select-Object $Select
            }
        } #end foreach name

    }
    End {
        Write-Progress * -Completed
        Write-Verbose "[END: $ScriptName]"
    }
} #end Resolve-PKDNSName

$Null = New-Alias -Name Resolve-PKDNS -Value Resolve-PKDNSName -Force -ErrorAction SilentlyContinue


