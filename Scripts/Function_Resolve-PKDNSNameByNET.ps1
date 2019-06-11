#requires -Version 3
Function Resolve-PKDNSNameByNET {
<#
.SYNOPSIS
    Uses [System.Net.Dns]::GetHostEntryAsync to resolves a hostname to an IP address, or an IP address to a hostname

.DESCRIPTION
    Uses [System.Net.Dns]::GetHostEntryAsync to resolves a hostname to an IP address, or an IP address to a hostname
    Works where the DNSClient module is not available
    Uses locally configured nameservers only
    Tests syntax of input using regex
    Accepts pipeline input
    Returns a PSObject

.NOTES
    Name    : Function_Resolve-PKDNSNameByNET.ps1
    Author  : Paula Kingsley
    Created : 2018-01-26
    Version : 03.00.0000    
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2018-01-26 - Created script
        v02.00.0000 - 2019-03-28 - Totally overhauled & renamed from Test-PKDNSResolution; removed ReturnTrueOnly
        v03.00.0000 - 2019-06-10 - Renamed to Resolve-PKDNSNameByNET. More overhaulage. Renamed InputObj to Name, 
                                   other edits/rearrangements, added warning about function limitations

.PARAMETER InputObj
    Name, FQDN, or IP address to look up

.PARAMETER Quiet
    Suppress all non-verbose console output

.EXAMPLE
    PS C:\> Resolve-PKDNSNameByNET -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value            
        ---           -----            
        Verbose       True             
        Name          {LAPTOP8}
        Quiet         False            
        PipelineInput False            
        ScriptName    Resolve-PKDNSNameByNET
        ScriptVersion 3.0.0            

        WARNING: This function performs a very simple DNS resolution test using the [System.Net.Dns]::GetHostEntryAsync() method.
        You cannot specify nameservers or perform specific lookup queries.
        For more options try Get-PKDNSResolution (which uses Resolve-DNSName).

        BEGIN  : Perform basic DNS hostname/IP resolution test using .NET method

        [LAPTOP8] Resolved hostname to IP address in 1 millisecond(s)

        Name       : LAPTOP8
        IsResolved : True
        ResolvesTo : {192.168.78.34, ::1}
        Nameserver : {127.0.0.1}
        Messages   : Resolved hostname to IP address in 1 millisecond(s)

        END    : Perform basic DNS hostname/IP resolution test using .NET method

.EXAMPLE
    PS C:\> Resolve-PKDNSNameByNET -Name 127.0.0.1,sqldev1.domain.local,google.com,!@#$%.com,thereisnowayanyoneregisteredthisdomain.info  -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                                                                         
        ---           -----                                                                         
        Name          {127.0.0.1, sqldev1.domain.local, google.com, !@#$%.com...}
        Verbose       True                                                                          
        Quiet         False                                                                         
        PipelineInput False                                                                         
        ScriptName    Resolve-PKDNSNameByNET                                                             
        ScriptVersion 3.0.0                                                                         


        WARNING: This function performs a very simple DNS resolution test using the [System.Net.Dns]::GetHostEntryAsync() method.
        You cannot specify nameservers or perform specific lookup queries.
        For more features try Resolve-PKDNSName (which uses Resolve-DNSName).

        BEGIN  : Perform basic DNS hostname/IP resolution test using .NET method

        [127.0.0.1] Resolved IPv4 address to hostname in 1 millisecond(s)
        Input      : 127.0.0.1
        IsResolved : True
        Output     : vmware-localhost
        Nameserver : {172.20.2.30, 172.21.2.14}
        Messages   : Resolved IPv4 address to hostname in 1 millisecond(s)

        [sqldev1.domain.local] Resolved hostname to IP address in 2 millisecond(s)
        Input      : sqldev1.domain.local
        IsResolved : True
        Output     : 10.62.179.193
        Nameserver : {172.20.2.30, 172.21.2.14}
        Messages   : Resolved hostname to IP address in 2 millisecond(s)

        [google.com] Resolved hostname to IP address in 2 millisecond(s)
        Input      : google.com
        IsResolved : True
        Output     : 216.58.199.174
        Nameserver : {172.20.2.30, 172.21.2.14}
        Messages   : Resolved hostname to IP address in 2 millisecond(s)

        [!@#$%.com] Invalid hostname syntax
        Input      : !@#$%.com
        IsResolved : False
        Output     : Error
        Nameserver : {172.20.2.30, 172.21.2.14}
        Messages   : Invalid hostname syntax

        [thereisnowayanyoneregisteredthisdomain.info] Failed to resolve hostname to IP address after 59 millisecond(s)
        Input      : thereisnowayanyoneregisteredthisdomain.info
        IsResolved : False
        Output     : Error
        Nameserver : {172.20.2.30, 172.21.2.14}
        Messages   : Failed to resolve hostname to IP address after 59 millisecond(s)

        END    : Perform basic DNS hostname/IP resolution test using .NET method


.EXAMPLE
    PS C:\> Get-VM pk* | Resolve-PKDNSNameByNET -Quiet -Warning:$False | Format-Table -AutoSize

        Name                IsResolved ResolvesTo    Nameserver                                    Messages                                                          
        ----                ---------- ----------    ----------                                    --------                                                          
        pk-prometheus-102        True 172.20.64.20  {10.10.45.30, 10.11.45.30} Resolved hostname to IP address in 1 millisecond(s)               
        pk-prometheus-101        True 172.20.64.19  {10.10.45.30, 10.11.45.30} Resolved hostname to IP address in 1 millisecond(s)               
        pk-prometheus-dev-1      True 172.20.64.24  {10.10.45.30, 10.11.45.30} Resolved hostname to IP address in 1 millisecond(s)               
        pk-test-501              True 10.52.128.255 {10.10.45.30, 10.11.45.30} Resolved hostname to IP address in 1 millisecond(s)               
        pk-prometheus-qa1       False Error         {10.10.45.30, 10.11.45.30} Failed to resolve hostname to IP address after 838 millisecond(s) 
        pk-prometheus-qa2       False Error         {10.10.45.30, 10.11.45.30} Failed to resolve hostname to IP address after 837 millisecond(s) 
        pk-prometheus-qa3       False Error         {10.10.45.30, 10.11.45.30} Failed to resolve hostname to IP address after 837 millisecond(s) 
        pkweb-1                  True 10.11.178.103 {10.10.45.30, 10.11.45.30} Resolved hostname to IP address in 1 millisecond(s)               
        pk-prometheus-1         False Error         {10.10.45.30, 10.11.45.30} Failed to resolve hostname to IP address after 838 millisecond(s) 
        pk-sql-1                 True 10.62.179.204 {10.10.45.30, 10.11.45.30} Resolved hostname to IP address in 1 millisecond(s)               
        pktestbox                True 10.11.129.166 {10.10.45.30, 10.11.45.30} Resolved hostname to IP address in 12 millisecond(s)              


#>
[CmdletBinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Hostname or IP for lookup"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$Name = $Env:ComputerName,

    [Parameter(
        HelpMessage = "Suppress all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "03.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    If ($PipelineInput.IsPresent) {$CurrentParams.InputObject = $Null}
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Make sure we know the limitations
    $Msg = "This function performs a very simple DNS resolution test using the [System.Net.Dns]::GetHostEntryAsync() method.`nYou cannot specify nameservers or perform specific lookup queries.`nFor more options try Resolve-PKDNSName (which uses Resolve-DNSName).`n"
    Write-Warning $Msg

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Write-Progress
    $Activity = "Perform basic DNS hostname/IP resolution test using .NET method"
    $Param_WP = @{
        Activity         = $Activity
        Status           = "Working"
        CurrentOperation = $Null
        PercentComplete  = $Null
    }

    #endregion Splats

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

    #endregion Functions

    #region Regex
    $InvalidNameRegex = '[\\%\/:"*?<>@)(~#&$|]+'
    $IPRegex = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$" # This is to separate IP format from hostname/string
    $ValidIPRegex = "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"

    # Get local DNS servers for output object
    $DNSServers = (Get-WmiObject -Namespace root\cimv2 -Query "Select dnsserversearchorder FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True" | Where-Object {$_.DNSServerSEarchOrder -ne $null} | Select-Object -ExpandProperty DNSServerSearchOrder)

    #endregion Regex
    
    # Console output
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}

Process {
    
    # Remove any dupes
    $Name = ($Name | Select-Object -Unique)

    $Total = ($Name -as [array]).Count
    $Current = 0

    Foreach ($Item in $Name) {

        # Create new object with additional properties
        $Results = ($Item | 
            Select-Object @{N="Name";E={$_.Trim()}},
            @{N="IsResolved";E={"Error"}},
            @{N="ResolvesTo";E={"Error"}},
            @{N="Nameserver";E={$DNSServers}},
            @{N="Messages";  E={"Error"}}
        )
    
        $Current ++
        $Param_WP.CurrentOperation = $Results.Input
        $Param_WP.PercentComplete = ($Current/$Total * 100)
        $Param_WP.Status = "$([math]::Round($Current/$Total * 100))%"
        Write-Progress @Param_WP

        # Make sure input is valid
        [switch]$Continue = $False
        If ($Results.Name -match $IPRegex) {
            If ($Results.Name -match $ValidIPRegex) {
                $Type = "IPv4 address"
                $ToType = "hostname"
                $Continue = $True
            }
            Else {
                $Msg = "Invalid IPv4 address syntax"
                "[$($Results.Name)] $Msg" | Write-MessageError
                $Results.Messages = $Msg
                $Results.IsResolved = $False

                Write-Output $Results
            }
        }
        Else {
            If ($Results.Name -notmatch $InvalidNameRegex) {
                $Type = "hostname"
                $ToType = "IP address"
                $Continue = $True
            }
            Else {
                $Msg = "Invalid hostname syntax"
                "[$($Results.Name)] $Msg" | Write-MessageError
                $Results.Messages = $Msg
                $Results.IsResolved = $False

                Write-Output $Results
            }
        } #end switch

        If ($Continue.IsPresent) {

            Try {
                
                # Start the timer
                $Stopwatch =  [system.diagnostics.stopwatch]::StartNew()
                $Task = $Null
                $Task = [System.Net.Dns]::GetHostEntryAsync($Results.Name)
                
                If ($Task.Result) {

                    $Stopwatch.Stop()    
                    $Elapsed = [math]::Round($Stopwatch.Elapsed.TotalMilliseconds)
                    $Msg = "Resolved $Type to $ToType in $Elapsed millisecond(s)"
                    "[$($Results.Name)] $Msg" | Write-MessageInfo -FGColor Green
                    
                    $Results.IsResolved = $True
                    $Results.ResolvesTo = Switch ($Type) {
                        "Hostname"     {$Task.Result.AddressList.IPAddressToString}    
                        "IPv4 address" {$Task.Result.Hostname}
                        Default        {}
                    }
                    $Results.Messages = $Msg

                    Write-Output $Results

                } #end if results
                Else {
                    $Stopwatch.Stop()
                    $Elapsed = [math]::Round($Stopwatch.Elapsed.TotalMilliseconds)
                    $Msg = "Failed to resolve $Type to $ToType after $Elapsed millisecond(s)"
                    "[$($Results.Name)] $Msg" | Write-MessageError

                    $Results.IsResolved = $False
                    $Results.Messages = $Msg

                    Write-Output $Results
                    
                } #end if no results

            }
            Catch {
                $Msg = "Operation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "$Msg" | Write-MessageError

                $Results.IsResolved = $False
                $Results.Messages = $Msg

                Write-Output $Results
            } 
        
            # Clean up
            $Task.Dispose()
            [system.gc]::Collect()
        }

        # Output
        #Write-Output $Results

    } #end Foreach
    
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title
    
}
} #end Resolve-PKDNSNameByNET

#$Null = New-Alias -Name Resolve-PKDNSRecord -Value Resolve-PKDNSNameByNET -Description "Guessability" -Force -Confirm:$False
#$Null = New-Alias -Name Resolve-PKDNSAddress -Value Resolve-PKDNSNameByNET -Description "Guessability" -Force -Confirm:$False
#$Null = New-Alias -Name Resolve-PKDNSIP -Value Resolve-PKDNSNameByNET -Description "Guessability" -Force -Confirm:$False

