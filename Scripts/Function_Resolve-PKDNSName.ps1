#requires -Version 3
Function Resolve-PKDNSName {
<#
.SYNOPSIS
    Uses [System.Net.Dns]::GetHostEntryAsync to resolves a name to an IP address, or an IP address to a hostname

.DESCRIPTION
    Uses [System.Net.Dns]::GetHostEntryAsync to resolves a name to an IP address, or an IP address to a hostname
    Uses locally configured nameservers (cannot be specified)
    Optional -ReturnTrueOnly switch does not return objects for failed queries
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
        v02.00.0000 - 2019-03-28 - Overhauled & renamed from Test-PKDNSResolution

.PARAMETER InputObj
    Name, FQDN, or IP address to look up

.PARAMETER SuppressConsoleOutput
    Suppress all non-verbose/non-error console output

.PARAMETER ReturnTrueOnly
    Return results only for entries that successfully resolve

.EXAMPLE
    PS C:\> Test-PKDNSResolution -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value            
        ---            -----            
        Verbose        True             
        Name           {WORKSTATION15}
        ReturnTrueOnly False            
        Quiet          False            
        ScriptName     Resolve-PKDNSName
        ScriptVersion  2.0.0            

        BEGIN  : Test DNS name/IP resolution

        [WORKSTATION15] Look up DNS record
        [WORKSTATION15] Name resolution succeeded

        END    : Test DNS name/IP resolution

        Input           IsResolved Output               Messages                 
        -----           ---------- ------               --------                 
        WORKSTATION15       True   {10.32.150.213, ::1} Name resolution succeeded

.EXAMPLE
     PS C:\>  Resolve-PKDNSName 127.0.0.1,sqlserver.domain.local,google.com,!@#$%.com,thereisnowayanyoneregisteredthisdomain.info -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                                                         
        ---            -----                                                                         
        Verbose        True                                                                          
        Name           {127.0.0.1, sqlserver.domain.local, google.com, !@#$%.com...}
        ReturnTrueOnly False                                                                         
        Quiet          False                                                                         
        ScriptName     Resolve-PKDNSName                                                             
        ScriptVersion  2.0.0                                                                         

        BEGIN  : Test DNS name/IP resolution

        [127.0.0.1] Look up DNS record
        [127.0.0.1] Name resolution succeeded

        [sqlserver.domain.local] Look up DNS record
        [sqlserver.domain.local] Name resolution succeeded
        [google.com] Look up DNS record
        [google.com] Name resolution succeeded
        [!@#$%.com] Look up DNS record
        [!@#$%.com] Invalid syntax
        [thereisnowayanyoneregisteredthisdomain.info] Look up DNS record
        [thereisnowayanyoneregisteredthisdomain.info] Name resolution failed
        
        END    : Test DNS name/IP resolution

        Input                                       IsResolved Output           Messages                 
        -----                                       ---------- ------           --------                 
        127.0.0.1                                         True vmware-localhost Name resolution succeeded
        sqlserver.domain.local                            True 10.11.178.173    Name resolution succeeded
        google.com                                        True 216.58.194.174   Name resolution succeeded
        !@#$%.com                                        False Error            Invalid syntax           
        thereisnowayanyoneregisteredthisdomain.info      False Error            Name resolution failed   

.EXAMPLE
    PS C:\> Resolve-PKDNSName 127.0.0.1,sqlserver.domain.local,google.com,!@#$%.com,thereisnowayanyoneregisteredthisdomain.info -Quiet -ReturnTrueOnly

        Input                  IsResolved Output           Messages                 
        -----                  ---------- ------           --------                 
        127.0.0.1                    True vmware-localhost Name resolution succeeded
        sqlserver.domain.local       True 10.11.178.173    Name resolution succeeded
        google.com                   True 216.58.194.174   Name resolution succeeded



#>
[CmdletBinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        HelpMessage = "Hostname or IP for lookup"
    )]
    [Alias("InputObj","IP","DNSHostName","FQDN")]
    [ValidateNotNullOrEmpty()]
    [object[]]$Name = $Env:ComputerName,

    [Parameter(
        HelpMessage = "Return output for successful name resolution only"
    )]
    [switch]$ReturnTrueOnly,

    [Parameter(
        HelpMessage = "Suppress all non-verbose console output"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"

    # Show our settings
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
    $ValidIPRegex = "\b(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}\b"

    #endregion Regex
    
    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg`n")}
    Else {Write-Verbose $Msg}

}

Process {
    
    # Remove any dupes
    $Name = ($Name | Select-Object -Unique)

    $Total = $Name.Count
    $Current = 0

    # Create new object with additional properties
    [array]$Lookup = ($Name | 
        Select-Object @{N="Input";E={$_}},
        @{N="IsResolved";E={"Error"}},
        @{N="Output";E={"Error"}},
        #@{N="NameServer";E={@(Get-WmiObject Win32_NetworkAdapterConfiguration | Select-Object -ExpandProperty IPAddress)}},
        @{N="Messages";E={"Error"}}
    )

    Try {
    
        Foreach ($Item in $Lookup) {
            
            $Current ++

            $Msg = "Look up DNS record"
            $FGColor = "White"
            If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$($Item.Input)] $Msg")}
            Else {Write-Verbose "[$($Item.Input)] $Msg"}

            $Param_WP.CurrentOperation = $Msg
            $Param_WP.PercentComplete = ($Current/$Total * 100)
            $Param_WP.Status = "$([math]::Round($Current/$Total * 100))%"
            Write-Progress @Param_WP
            
            # Object to return
            $Output = @() 
            
             If ($Item.Input -match $InvalidNameRegex) {
                $Msg = "Invalid syntax"
                If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("[$($Item.Input)] $Msg")}
                Else {Write-Verbose $Msg}
                
                If (-not $ReturnTrueOnly.IsPresent) {
                    $Item.IsResolved = $False
                    $Item.Messages = $Msg
                    $Output = $Item
                }
             }
            
            Else {
                $Task = $Null
                $Task = [System.Net.Dns]::GetHostEntryAsync($Item.Input)

                If ($Task.Result) {
                
                    $Msg = "Name resolution succeeded"
                    $FGColor = "Green"
                    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"[$($Item.Input)] $Msg")}
                    Else {Write-Verbose "[$($Item.Input)] $Msg"}

                    $Item.IsResolved = $True
                    $Item.Output = Switch -Regex ($Item.Input) {
                        $ValidIPRegex {$Task.Result.Hostname}
                        Default       {$Task.Result.AddressList.IPAddressToString}
                    }
                    $Item.Messages = $Msg

                    $Output = $Item
                }
                Else {
                    $Msg = "Name resolution failed"
                    If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("[$($Item.Input)] $Msg")}
                    Else {Write-Verbose $Msg}

                    If (-not $ReturnTrueOnly.IsPresent) {
                        $Item.IsResolved = $False
                        $Item.Messages = $Msg
                        $Output = $Item
                    }
                }

                # Clean up
                $Task.Dispose()
            }

            Write-Output $Output

        } #end for each
    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
        If (-not $Quiet.IsPresent) {$Host.UI.ErrorLine($Msg)}
        Else {Write-Verbose $Msg}

        If (-not $ReturnTrueOnly.IsPresent) {
            New-Object PSObject -Property ([ordered]@{
                Input      = $Name
                IsResolved = $False
                Output     = "Error"
                #NameServer = (Get-WmiObject Win32_NetworkAdapterConfiguration | Select-Object -ExpandProperty IPAddress)
                Messages   = $Msg
            })
        }
    }


}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    
    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg`n")}
    Else {Write-Verbose $Msg}
    
}
} #end Resolve-PKDNSName

$Null = New-Alias -Name Resolve-PKDNSRecord -Value Resolve-PKDNSName -Description "Guessability" -Force -Confirm:$False



<#
 # Function if we want to check additional nameservers
    Function LookupDNS { 
        [Cmdletbinding()]
        Param(
            [string[]]$Target,
            $NS
        )
        [switch]$NSLookup = $False
        If ($NS) {
            Write-Verbose "Nameserver: $NS"
            If ($TestNS = (new-object Net.Sockets.TcpClient -ErrorAction Stop -Verbose)) {
                If (-not ($TestNS.ConnectAsync($NS,53))) {
                    $Msg = "No response from nameserver $NS"
                    $Host.UI.WriteErrorLine($Msg)
                    Break
                }
            }
            Else {$NSLookup = $True}
        }
        Else {
            Write-Verbose "Nameserver: Local default"
        }

        Foreach ($T in $Target) {

            $Output = New-Object PSObject -Property ([ordered]@{
                Input      = $T
                Mode       = "Error"
                Name       = "Error"
                IPAddress  = "Error"
                Messages   = "Error"
            })

            # If IP address
            If ($T -as [ipaddress]) {
                
                $Output.Mode = "Reverse"

                If ($NSLookup.IsPresent) {

                    If ($Lookup = Invoke-Expression -Command "nslookup $T $NS 2>&1" -OutVariable Results -EA SilentlyContinue) {
                        $Output.IPAddress = ($Results | Select-String address | select-Object -last 1).toString().split(":")[1].trim()
                        $Output.Name = ($Results | Select-String name | select-Object -last 1).toString().split(":")[1].trim()
                        $Output.Messages = $Null
                    }
                    Else {$Output.Messages = "Reverse lookup failed for $T"}
                }
                Else {
                    If ($Results = [System.Net.DNS]::GetHostByAddress($T)) {
                        $Output.IPAddress = $($Results.AddressList)                
                        $Output.Name = ($Results.hostname.tolower())
                        $Output.Messages = $Null
                    }
                    Else {$Output.Messages = "Reverse lookup failed for $T"}
                }
            }
            
            # If hostname 
            Else {

                $Output.Mode = "Forward"

                If ($NSLookup.IsPresent) {
                    
                    If ($Lookup = Invoke-Expression -Command "nslookup $T $NS 2>&1" -OutVariable Results -EA SilentlyContinue) {
                        $Output.IPAddress = ($Results | Select-String address | select-Object -last 1).toString().split(":")[1].trim()
                        $Output.Name = ($Results | Select-String name | select-Object -last 1).toString().split(":")[1].trim()
                        $Output.Messages = $Null
                    }
                    Else {$Output.Messages = "Forward lookup failed for $T"}
                }
                Else {
                    If ($Results = [System.Net.DNS]::GetHostByName($T)) {
                        $Output.IPAddress = $($Results.AddressList)                
                        $Output.Name = ($Results.hostname.tolower())
                        $Output.Messages = $Null
                    }
                    Else {$Output.Messages = "Forward lookup failed for $T"}
                }
            }
            
            Write-Output $Output
        }
    }
    
    #endregion Functions
#>