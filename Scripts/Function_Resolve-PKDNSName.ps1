#requires -Version 3
Function Resolve-PKDNSName {
<#
.SYNOPSIS
    Uses [System.Net.Dns]::GetHostEntryAsync to resolves a hostname to an IP address, or an IP address to a hostname

.DESCRIPTION
    Uses [System.Net.Dns]::GetHostEntryAsync to resolves a hostname to an IP address, or an IP address to a hostname
    Works where the DNSClient module is not available
    Uses locally configured nameservers only
    Accepts pipeline input
    Returns a PSObject

    NOTE: For a more fully-functional DNS test using Resolve-DNSName, try Test-PKDNSServer

.NOTES
    Name    : Function_Resolve-PKDNSName.ps1
    Author  : Paula Kingsley
    Created : 2018-01-26
    Version : 02.00.0000    
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2018-01-26 - Created script
        v02.00.0000 - 2019-03-28 - Totally overhauled & renamed from Test-PKDNSResolution; removed ReturnTrueOnly

.PARAMETER InputObj
    Name, FQDN, or IP address to look up

.PARAMETER Quiet
    Suppress all non-verbose console output

.EXAMPLE
    PS C:\> Resolve-PKDNSName -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value            
        ---            -----            
        Verbose        True             
        Name           {LAPTOP14}
        ReturnTrueOnly False            
        Quiet          False            
        ScriptName     Resolve-PKDNSName
        ScriptVersion  2.0.0            
        PipelineInput  False            

        BEGIN  : Test DNS name/IP resolution
        [LAPTOP14] Resolved hostname to IP address in 0 milliseconds
        
        Input      : LAPTOP14
        IsResolved : True
        Output     : {172.20.2.30, ::1}
        Nameserver : 172.20.2.30
        Messages   : Resolved hostname to IP address in 0 milliseconds

        END    : Test DNS name/IP resolution

.EXAMPLE
     PS C:\>  Resolve-PKDNSName 127.0.0.1,sqlserver.domain.loca,google.com,!@#$%.com,thereisnowayanyoneregisteredthisdomain.info 

        BEGIN  : Test DNS name/IP resolution
        [127.0.0.1] Resolved IPv4 address to hostname in 1 milliseconds

        Input      : 127.0.0.1
        IsResolved : True
        Output     : vmware-localhost
        Nameserver : 172.20.2.30
        Messages   : Resolved IPv4 address to hostname in 1 milliseconds

        [sqlserver.domain.loca] Resolved hostname to IP address in 2 milliseconds
        Input      : sqlserver.domain.loca
        IsResolved : True
        Output     : 10.62.179.193
        Nameserver : 172.20.2.30
        Messages   : Resolved hostname to IP address in 2 milliseconds

        [google.com] Resolved hostname to IP address in 2 milliseconds
        Input      : google.com
        IsResolved : True
        Output     : 172.217.3.206
        Nameserver : 172.20.2.30
        Messages   : Resolved hostname to IP address in 2 milliseconds

        [!@#$%.com] Invalid hostname syntax
        Input      : !@#$%.com
        IsResolved : False
        Output     : Error
        Nameserver : 172.20.2.30
        Messages   : Invalid hostname syntax

        [thereisnowayanyoneregisteredthisdomain.info] Failed to resolve hostname to IP address after 37 milliseconds
        Input      : thereisnowayanyoneregisteredthisdomain.info
        IsResolved : False
        Output     : Error
        Nameserver : 172.20.2.30
        Messages   : Failed to resolve hostname to IP address after 37 milliseconds

        END    : Test DNS name/IP resolution

.EXAMPLE
    PS C:\> Get-AllMyDCs -ADDomain domain.local -All | Select -first 5 | Resolve-PKDNSName -Verbose | Format-Table -AutoSize
 
        Action: Get all Active Directory Domain Controllers in domain
        VERBOSE: PSBoundParameters: 
	
        Key            Value            
        ---            -----            
        Verbose        True             
        Name           
        ReturnTrueOnly False            
        Quiet          False            
        PipelineInput  True             
        ScriptName     Resolve-PKDNSName
        ScriptVersion  2.0.0            

        BEGIN  : Test DNS name/IP resolution
        [RomeDC03] Resolved hostname to IP address in 3 milliseconds
        [RomeDC01] Resolved hostname to IP address in 2 milliseconds
        [LondonDC01] Resolved hostname to IP address in 3 milliseconds
        [ParisDC02] Resolved hostname to IP address in 2 milliseconds
        [MadridDC02] Resolved hostname to IP address in 7 milliseconds
        END    : Test DNS name/IP resolution
        Input      IsResolved Output        Nameserver  Messages                                         
        -----      ---------- ------        ----------  --------                                         
        RomeDC03       True   10.50.142.103 172.20.2.30 Resolved hostname to IP address in 3 milliseconds
        RomeDC01       True   10.92.142.11  172.20.2.30 Resolved hostname to IP address in 2 milliseconds
        LondonDC01     True   10.44.136.43  172.20.2.30 Resolved hostname to IP address in 3 milliseconds
        ParisDC02      True   10.12.128.198 172.20.2.30 Resolved hostname to IP address in 2 milliseconds
        MadridDC02     True   10.76.136.30  172.20.2.30 Resolved hostname to IP address in 7 milliseconds

#>
[CmdletBinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        HelpMessage = "Hostname or IP for lookup"
    )]
    [Alias("Name","InputObj","IP","DNSHostName","FQDN","IPv4Address")]
    [ValidateNotNullOrEmpty()]
    $InputObject = $Env:ComputerName,

    [Parameter(
        HelpMessage = "Suppress all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"

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
    $Activity = "Test DNS name/IP resolution"
    If ($ReturnTrueOnly.IsPresent) {$Activity += " (return only valid results)"}
    $Param_WP = @{
        Activity         = $Activity
        Status           = "Working"
        CurrentOperation = $Null
        PercentComplete  = $Null
    }

    #endregion Splats

    #region Regex
    $InvalidNameRegex = '[\\%\/:"*?<>@)(~#&$|]+'
    $IPRegex = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$" # This is to separate IP format from hostname/string
    $ValidIPRegex = "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"

    # Get local DNS servers
    $DNSServers = (Get-WmiObject Win32_NetworkAdapterConfiguration | Select-Object -ExpandProperty IPAddress)

    #endregion Regex
    
    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}

}

Process {
    
    # Remove any dupes
    $InputObject = ($InputObject | Select-Object -Unique)

    $Total = ($InputObject -as [array]).Count
    $Current = 0

    # Create new object with additional properties
    [array]$Output = ($InputObject | 
        Select-Object @{N="Input";E={$_}},
        @{N="IsResolved";E={"Error"}},
        @{N="Output";E={"Error"}},
        @{N="Nameserver";E={$DNSServers}},
        @{N="Messages";E={"Error"}}
    )

    Try {
        Foreach ($Item in $Output) {
            
            $Current ++

            $Msg = "Look up DNS record"
            $Param_WP.CurrentOperation = $Msg
            $Param_WP.PercentComplete = ($Current/$Total * 100)
            $Param_WP.Status = "$([math]::Round($Current/$Total * 100))%"
            Write-Progress @Param_WP

            # Make sure input is valid
            [switch]$Continue = $False
            If ($Item.Input -match $IPRegex) {
                If ($Item.Input -match $ValidIPRegex) {
                    $Type = "IPv4 address"
                    $ToType = "hostname"
                    $Continue = $True
                }
                Else {
                    $Msg = "Invalid IPv4 address syntax"
                    If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("[$($Item.Input)] $Msg")}
                    Else {Write-Warning $Msg}
                    $Item.Messages = $Msg
                    $Item.IsResolved = $False
                }
            }
            Else {
                If ($Item.Input -notmatch $InvalidNameRegex) {
                    $Type = "hostname"
                    $ToType = "IP address"
                    $Continue = $True
                }
                Else {
                    $Msg = "Invalid hostname syntax"
                    If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("[$($Item.Input)] $Msg")}
                    Else {Write-Warning $Msg}
                    $Item.Messages = $Msg
                    $Item.IsResolved = $False
                }
            } #end switch

            If ($Continue.IsPresent) {
                
                # Start the timer
                $Stopwatch =  [system.diagnostics.stopwatch]::StartNew()
                $Task = $Null
                $Task = [System.Net.Dns]::GetHostEntryAsync($Item.Input)
                
                If ($Task.Result) {

                    $Stopwatch.Stop()    
                    $Elapsed = [math]::Round($Stopwatch.Elapsed.TotalMilliseconds)
                    $Msg = "Resolved $Type to $ToType in $Elapsed millisecond(s)"
                    $FGColor = "Green"
                    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$($Item.Input)] $Msg")}
                    Else {Write-Verbose "[$($Item.Input)] $Msg"}

                    $Item.IsResolved = $True

                    $Item.Output = Switch ($Type) {
                        "Hostname"   {
                            $Task.Result.AddressList.IPAddressToString
                        }    
                        "IPv4 address" {
                            $Task.Result.Hostname
                        }
                        Default {}
                    }
                    $Item.Messages = $Msg
                } #end if results
                Else {
                    $Stopwatch.Stop()
                    $Elapsed = [math]::Round($Stopwatch.Elapsed.TotalMilliseconds)
                    $Msg = "Failed to resolve $Type to $ToType after $Elapsed millisecond(s)"
                    If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("[$($Item.Input)] $Msg")}
                    Else {Write-Verbose $Msg}

                    If (-not $ReturnTrueOnly.IsPresent) {
                        $Item.IsResolved = $False
                        $Item.Messages = $Msg
                    }
                } #end if no results

                # Clean up
                $Task.Dispose()
            }

            Write-Output $Item

        } #end for each
    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine($Msg)}
        Else {Write-Verbose $Msg}

        If (-not $ReturnTrueOnly.IsPresent) {
            New-Object PSObject -Property ([ordered]@{
                Input      = $InputObject
                IsResolved = $False
                Output     = "Error"
                NameServer = $DNSServers
                Messages   = $Msg
            })
        }
    }


}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    
    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}
    
}
} #end Resolve-PKDNSName

$Null = New-Alias -Name Resolve-PKDNSRecord -Value Resolve-PKDNSName -Description "Guessability" -Force -Confirm:$False
$Null = New-Alias -Name Resolve-PKDNSAddress -Value Resolve-PKDNSName -Description "Guessability" -Force -Confirm:$False
$Null = New-Alias -Name Resolve-PKDNSIP -Value Resolve-PKDNSName -Description "Guessability" -Force -Confirm:$False

