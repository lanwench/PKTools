#requires -Version 3
Function Resolve-PKDNSName {
<#
.SYNOPSIS
    Uses Resolve-DNSName to test lookups on a DNS server, with options for record type and truncated output

.DESCRIPTION
    Uses Resolve-DNSName to test lookups on a DNS server, with options for record type and truncated output
    Defaults to microsoft.com as lookup target
    Defaults to locally-configured DNS servers
    By default attempts to first ping DNS server and connect via TCP on port 53, but individual tests can be selected or skipped outright
    Accepts pipeline input for names
    Returns a PSObject

.NOTES
    Name    : Function_Resolve-PKDNSName.ps1 
    Created : 2019-06-10
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK
        
        v01.00.0000 - 2019-06-10 - Created based on Test-PKDNSServer

.PARAMETER Server
    DNS server name or IP (default is locally configured DNS servers)

.PARAMETER Lookup
    Name or IP address for record lookup (default is local computername)

.PARAMETER RecordType
    DNS record type: Any, A, AAAA, CNAME, MX, NS, PTR, SOA, SRV, TXT (default is 'A')

.PARAMETER NoRecursion
    Don't perform recursive lookup

.PARAMETER TruncateLookupOutput
    Truncate lookup data (useful if only basic testing is needed and the full output is lengthy)
    
.PARAMETER Quiet
    Suppress non-verbose console output

.EXAMPLE
    PS C:\> Resolve-PKDNSName -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                  Value                                        
        ---                  -----                                        
        Verbose              True                                         
        Server               {172.32.4.250, 172.33.4.250, 172.34.4.250}
        Name                 {WORKSTATION14}                            
        RecordType           A                                            
        NoRecursion          False                                        
        TruncateLookupOutput False                                        
        Quiet                False                                        
        ScriptName           Resolve-PKDNSName                            
        ScriptVersion        1.0.0                                        

        BEGIN  : Resolve DNS name or IP address using Resolve-DNSName

        [WORKSTATION14] Server 172.32.4.250: Perform DNS lookup (record type: A)
        [WORKSTATION14] Server 172.32.4.250: Host address record lookup succeeded in 11 milliseconds
        [WORKSTATION14] Server 172.33.4.250: Perform DNS lookup (record type: A)
        [WORKSTATION14] Server 172.33.4.250: Host address record lookup succeeded in 4 milliseconds
        [WORKSTATION14] Server 172.34.4.250: Perform DNS lookup (record type: A)
        [WORKSTATION14] Server 172.34.4.250: Host address record lookup succeeded in 5 milliseconds


        Name         : WORKSTATION14
        IsResolved   : True
        RecordType   : A
        Recursive    : True
        Output       : @{Address=10.42.50.178; IPAddress=10.42.50.178; QueryType=A; IP4Address=10.42.50.178; 
                       Name=workstation14.domain.local; Type=A; CharacterSet=Unicode; Section=Question; 
                       DataLength=4; TTL=1200}
        Server       : 172.32.4.250
        ComputerName : WORKSTATION14
        SourceIP     : {10.42.50.178}
        Messages     : Host address record lookup succeeded in 11 milliseconds

        Name         : WORKSTATION14
        IsResolved   : True
        RecordType   : A
        Recursive    : True
        Output       : @{Address=10.42.50.178; IPAddress=10.42.50.178; QueryType=A; IP4Address=10.42.50.178; 
                       Name=workstation14.domain.local; Type=A; CharacterSet=Unicode; Section=Question; 
                       DataLength=4; TTL=1200}
        Server       : 172.33.4.250
        ComputerName : WORKSTATION14
        SourceIP     : {10.42.50.178}
        Messages     : Host address record lookup succeeded in 4 milliseconds

        Name         : WORKSTATION14
        IsResolved   : True
        RecordType   : A
        Recursive    : True
        Output       : @{Address=10.42.50.178; IPAddress=10.42.50.178; QueryType=A; IP4Address=10.42.50.178; 
                       Name=workstation14.domain.local; Type=A; CharacterSet=Unicode; Section=Question; 
                       DataLength=4; TTL=1200}
        Server       : 172.34.4.250
        ComputerName : WORKSTATION14
        SourceIP     : {10.42.50.178}
        Messages     : Host address record lookup succeeded in 5 milliseconds

        END    : Resolve DNS name or IP address using Resolve-DNSName

.EXAMPLE
    PS C:\> Resolve-PKDNSName google.com -Server 172.20.2.22 -RecordType SOA -NoRecursion

        BEGIN  : Resolve DNS name or IP address using Resolve-DNSName

        [google.com] Server 172.20.2.22: Perform non-recursive DNS lookup (record type: SOA)
        [google.com] Server 172.20.2.22: Start of Authority record lookup succeeded in 23 milliseconds

        Name         : google.com
        IsResolved   : True
        RecordType   : SOA
        Recursive    : False
        Output       : @{QueryType=SOA; Administrator=dns-admin.google.com; PrimaryServer=ns1.google.com; 
                       NameAdministrator=dns-admin.google.com; SerialNumber=252617051; TimeToZoneRefresh=900; 
                       TimeToZoneFailureRetry=900; TimeToExpiration=1800; DefaultTTL=60; Name=google.com; Type=SOA; 
                       CharacterSet=Unicode; Section=Answer; DataLength=112; TTL=55}
        Server       : 172.20.2.22
        ComputerName : LAPTOP
        SourceIP     : {198.6.55.32}
        Messages     : Start of Authority record lookup succeeded in 23 milliseconds


        END    : Resolve DNS name or IP address using Resolve-DNSName

.EXAMPLE
    PS C:\> Resolve-PKDNSName microsoft.com -Server 198.6.75.200 -RecordType Any -NoRecursion -Quiet

        Name         : microsoft.com
        IsResolved   : True
        RecordType   : Any
        Recursive    : False
        Output       : {@{Address=40.76.4.15; IPAddress=40.76.4.15; QueryType=A; IP4Address=40.76.4.15; 
                       Name=microsoft.com; Type=A; CharacterSet=Unicode; Section=Answer; DataLength=4; TTL=3600}, 
                       @{Address=13.77.161.179; IPAddress=13.77.161.179; QueryType=A; IP4Address=13.77.161.179; 
                       Name=microsoft.com; Type=A; CharacterSet=Unicode; Section=Answer; DataLength=4; TTL=3600}, 
                       @{Address=40.112.72.205; IPAddress=40.112.72.205; QueryType=A; IP4Address=40.112.72.205; 
                       Name=microsoft.com; Type=A; CharacterSet=Unicode; Section=Answer; DataLength=4; TTL=3600}, 
                       @{Address=104.215.148.63; IPAddress=104.215.148.63; QueryType=A; IP4Address=104.215.148.63; 
                       Name=microsoft.com; Type=A; CharacterSet=Unicode; Section=Answer; DataLength=4; TTL=3600}...}
        Server       : 198.6.75.200
        ComputerName : LAPTOP
        SourceIP     : {198.6.55.32}
        Messages     : DNS record (any type) lookup succeeded in 67 milliseconds

.EXAMPLE
    PS C:\> Resolve-PKDNSName microsoft.com -Server 1.1.1.1,8.8.8.8 -RecordType Any -NoRecursion -TruncateLookupOutput -Quiet

        Name         : microsoft.com
        IsResolved   : True
        RecordType   : Any
        Recursive    : False
        Output       : @{IP4Address=40.112.72.205; QueryType=A; Name=microsoft.com; Type=A}
        Server       : 1.1.1.1
        ComputerName : LAPTOP
        SourceIP     : {198.6.55.32}
        Messages     : DNS record (any type) lookup succeeded in 64 milliseconds

        Name         : microsoft.com
        IsResolved   : True
        RecordType   : Any
        Recursive    : False
        Output       : @{IP4Address=13.77.161.179; QueryType=A; Name=microsoft.com; Type=A}
        Server       : 8.8.8.8
        ComputerName : LAPTOP
        SourceIP     : {198.6.55.32}
        Messages     : DNS record (any type) lookup succeeded in 66 milliseconds

.EXAMPLE
    PS C:\> Resolve-PKDNSName thereisnowaythisdomainhaseverbeenregistered.info,blah^^^kitten,microsoft.com -RecordType MX -Server 198.6.75.200 -Quiet

        Name         : thereisnowaythisdomainhaseverbeenregistered.info
        IsResolved   : False
        RecordType   : MX
        Recursive    : True
        Output       : thereisnowaythisdomainhaseverbeenregistered.info : DNS name does not exist
        Server       : 198.6.75.200
        ComputerName : LAPTOP
        SourceIP     : {198.6.55.32}
        Messages     : Mail eXchange record lookup failed after 41 milliseconds

        Name         : blah^^^kitten
        IsResolved   : False
        RecordType   : MX
        Recursive    : True
        Output       : blah^^^kitten : DNS name contains an invalid character
        Server       : 198.6.75.200
        ComputerName : LAPTOP
        SourceIP     : {198.6.55.32}
        Messages     : Mail eXchange record lookup failed after 15 milliseconds

        Name         : microsoft.com
        IsResolved   : True
        RecordType   : MX
        Recursive    : True
        Output       : @{QueryType=MX; Exchange=microsoft-com.mail.protection.outlook.com; 
                       NameExchange=microsoft-com.mail.protection.outlook.com; Preference=10; Name=microsoft.com; 
                       Type=MX; CharacterSet=Unicode; Section=Answer; DataLength=100; TTL=3600}
        Server       : 198.6.75.200
        ComputerName : LAPTOP
        SourceIP     : {198.6.55.32}
        Messages     : Mail eXchange record lookup succeeded in 47 milliseconds


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
        HelpMessage = "Name or IP target for DNS lookup (default is local computer name)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Target","Lookup")]
    [string[]]$Name = $Env:ComputerName,

    [Parameter(
        Position = 1,
        HelpMessage = "One or more DNS server names or IP addresses"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Nameserver","DNSServer","IPAddress","IPv4Address","HostName")]
    [string[]]$Server,

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
        HelpMessage = "Truncate lookup data (useful if only basic testing is needed and the full output is lengthy)"
    )]
    [switch]$TruncateLookupOutput,

    [Parameter(
        HelpMessage = "Suppress non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch]$Quiet

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"
    
    $CurrentNameParams = $PSBoundParameters
    If (-not $CurrentNameParams.Server) { 
        $CurrentNameParams.Server = $Server = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE | Select-Object -Expand DNSServerSearchOrder)
    }
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentNameParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentNameParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentNameParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentNameParams.Add("ScriptVersion",$Version)
    
    Write-Verbose "PSBoundParameters: `n`t$($CurrentNameParams | Format-Table -AutoSize | Out-String )"

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

    # Make sure we know the limitations
    $Msg = "This function performs a DNS lookup using the Resolve-DNSName cmdlet (part of the DNSClient module).`nThe function Test-PKDNSServer adds the option to test connectivity and return Boolean output for the query.`nThe function Resolve-PKDNSNameByNET uses the [System.Net.Dns]::GetHostEntryAsync() method for basic lookups on systems without the DNSClient module.`n"
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
    
    #region Output object

    [switch]$IsRecursive = (-not $NoRecursion)
    $InitialValue = "Error"
    $OutputTemplate = [PSCustomObject]@{
        Name           = $InitialValue
        IsResolved     = $InitialValue
        RecordType     = $RecordType
        Recursive      = $IsRecursive.IsPresent
        Output         = $InitialValue
        Server         = $Server
        ComputerName   = $Env:ComputerName
        SourceIP       = @(@(Get-CIMInstance -Class Win32_NetworkAdapterConfiguration -Verbose:$False -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress) -like "*.*") #-join(", ")
        Messages       = $InitialValue
    }

    # Hashtable for record type description
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

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Test-Lookup
    $Param_Lookup = @{}
    $Param_Lookup = @{
        DNSServer     = $Null
        Target        = $Null
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
    $Activity = "Resolve DNS name or IP address using Resolve-DNSName"
    $Param_WP1 = @{}
    $Param_WP1 = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = $Null
        PercentComplete  = $Null
    }

    # Inner write-progress
    $Activity2 = "Perform DNS lookup (record type: $RecordType)"
    If ($NoRecursion.IsPresent) {$Activity2 = "Perform non-recursive DNS lookup (record type: $RecordType)"}
    $Param_WP2 = @{}
    $Param_WP2 = @{
        Activity         = $Activity
        ID               = 1
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    #endregion Splats

    # Console output
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}
Process {

    
    $TotalNames = $Name.Count
    $CurrentName = 0

    Foreach ($N in $Name) {
        
        $TotalServers = $Server.Count
        $CurrentServer = 0

        $CurrentName ++
        $Param_WP1.Status = $N
        $Param_WP1.PercentComplete = ($CurrentName/$TotalNames*100)
        $Param_Lookup.Target = $N

        $Results = @()
        [switch]$Continue = $False

        $Output1 = $OutputTemplate.PSObject.Copy()
        $Output1.Name = $N

        # Make sure we aren't trying to look up an IP address as type CNAME, for example
        If (-not (Test-CompatibleType -RecordType $RecordType -Name $N -ErrorAction Stop)) {
            $Msg = "Incompatible lookup target '$N' for DNS record type '$RecordType'"
            "[$N] $Msg" | Write-MessageError 
            $Output1.Messages = $Msg
            $Results += $Output1
        }

        # Just in case someone hasn't been careful with their parameters (ask me how I know)
        If ($N -in $Server) {
            $Msg = "DNS server and lookup name are identical"
            "[$N] $Msg" | Write-MessageInfo -FGColor Red
            $ConfirmMsg = "`n$Msg`nDo you wish to proceed with DNS query?`n"
            If ($PSCmdlet.ShouldContinue($ConfirmMsg,$Server)) {
                $Continue = $True
            }
            Else {
                $Msg = "Operation cancelled by user"
                "[$N] $Msg"  | Write-MessageInfo -FGColor Cyan
                $Output1.IsResolved = $False
                $Output1.Messages = $Msg
                $Results += $Output1
            }
        }
        Else {
            $Continue = $True
        }

        # Look up record
        If ($Continue.IsPresent) {
            
            Foreach ($S in $Server) {
            
                $CurrentServer ++

                $Param_WP1.CurrentOperation = $S
                Write-Progress @Param_WP1

                $Param_WP2.CurrentOperation = $Activity2
                $Param_WP2.Status = $S
                $Param_WP2.PercentComplete = ($CurrentServer/$TotalServers*100)
                Write-Progress @Param_WP2

                "[$N] Server $S`: $Activity2" | Write-MessageInfo -FGColor White
                
                Try {
                    
                    $Output2 = $Output1.PSObject.Copy()
                    $Output2.Server = $S

                    $Param_Lookup.DNSServer = $S
                    $StartTime = Get-Date
                    $Test = Test-Lookup @Param_Lookup
                    $EndTime = Get-Date
                    $Elapsed = ($EndTime - $StartTime).MilliSeconds
                    
                    <#
                    Switch -Regex ($Test) {
                        $False   {
                            $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                            "[$N] Server $S`: $Msg" | Write-MessageError
                            $Output2.IsResolved = $False
                            $Output2.Messages = $Msg
                        }
                        "FAIL: " {
                            $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                            "[$N] Server $S`: $Msg" | Write-MessageError
                            $Output2.IsResolved = $False
                            $Output2.Output = $Test.Replace("FAIL: ",$Null)
                            $Output2.Messages = $Msg
                        }
                        Default  {
                            $Msg = "$($RecordLookup[$RecordType]) lookup succeeded in $Elapsed milliseconds"
                            "[$N] Server $S`: $Msg"  | Write-MessageInfo -FGColor Green

                            $Output2.IsResolved = $True
                            If ($TruncateLookupOutput.IsPresent) {
                                $Output2.Output = ($Test | Select-Object -First 1) 
                            }
                            Else {
                                $Output2.Output = $Test
                            }
                            $Output2.Messages = $Msg   
                        }
                    }
                    #>

                    If ($Test -eq $False) {
                        $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                        "[$N] Server $S`: $Msg" | Write-MessageError
                        $Output2.IsResolved = $False
                        $Output2.Messages = $Msg
                    }
                    Elseif ($Test -match "FAIL: ") {
                        $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                        "[$N] Server $S`: $Msg" | Write-MessageError
                        $Output2.IsResolved = $False
                        $Output2.Output = $Test.Replace("FAIL: ",$Null)
                        $Output2.Messages = $Msg
                    }
                    Else  {
                        $Msg = "$($RecordLookup[$RecordType]) lookup succeeded in $Elapsed milliseconds"
                        "[$N] Server $S`: $Msg"  | Write-MessageInfo -FGColor Green

                        $Output2.IsResolved = $True
                        If ($TruncateLookupOutput.IsPresent) {
                            $Output2.Output = ($Test | Select-Object -First 1) 
                        }
                        Else {
                            $Output2.Output = $Test
                        }
                        $Output2.Messages = $Msg   
                    }
                }
                Catch {
                    $EndTime = Get-Date
                    $Elapsed = ($EndTime - $StartTime).MilliSeconds
                    $Msg = "$($RecordLookup[$RecordType]) lookup failed after $Elapsed milliseconds"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    "[$N] Server $S`: $Msg" | Write-MessageError
                    $Output2.Messages = $Msg
                    $Output2.IsResolved = $False
                }

                $Results += $Output2

            } # end for each server
        } # end if continue

        Write-Output $Results

    } # end for each name

}
End {
    
    Write-Progress -Activity $Activity -Completed
    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}    

} #end Resolve-PKDNSName


$Null = New-Alias Test-PKDNSResolution -Value Resolve-PKDNSName -Description "Guessability" -Force -Confirm:$False -ErrorAction SilentlyContinue
