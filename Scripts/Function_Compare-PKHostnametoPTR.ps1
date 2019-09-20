#requires -version 3
Function Compare-PKHostnameToPTR {
<#
.SYNOPSIS
    Compares a hostname to the PTR of its IP, using the .NET system.net.dns class

.DESCRIPTION
    Performs a reverse lookup on an IP address using the .NET system.net.dns class
    Accepts pipeline input
    Returns a PSObject or a boolean
     
.NOTES
    Name    : Function_Compare-PKHostnameToPTR.ps1
    Author  : Paula Kingsley
    Created : 2016-09-27
    Version : 01.00.0000
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2019-07-19 - Created script
                                                 
                              
.EXAMPLE
    PS C:\> Compare-PKHostnameToPTR -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                            
        ---            -----                            
        Verbose        True                             
        IPAddress      10.11.12.13                    
        Hostname       paula-laptop.domain.local
        ForceFQDNMatch False                            
        OutputType     Full                             
        Quiet          False                            
        PipelineInput  False                            
        ScriptName     ComparePKHostnameToPTR        
        ScriptVersion  2.0.0                            

        BEGIN: Match PTR to hostname

        [10.11.12.13] Look up PTR record
        [10.11.12.13] PTR/hostname match

        IPAddress     : 10.11.12.13
        Hostname      : paula-laptop.domain.local
        ReverseLookup : paula-laptop.domain.local
        IsMatch       : True
        DNSServers    : {10.60.157.250, 10.61.179.250, 10.62.179.250}
        Messages      : PTR/hostname match

        END  : Match PTR to hostname

.EXAMPLE
    PS C:\> Compare-PKHostnameToPTR -IPAddress 10.11.12.13 -HostName paula-laptop.blah.com -ForceFQDNMatch

        BEGIN: Match PTR to hostname (force FQDN match)

        [10.11.12.13] Look up PTR record
        [10.11.12.13] PTR/hostname mismatch

        IPAddress     : 10.11.12.13
        Hostname      : paula-laptop.blah.com
        ReverseLookup : paula-laptop.domain.local
        IsMatch       : False
        DNSServers    : {10.60.157.250, 10.61.179.250, 10.62.179.250}
        Messages      : PTR/hostname mismatch


        END  : Match PTR to hostname (force FQDN match)

.EXAMPLE
    PS C:\> Compare-PKHostnameToPTR -IPAddress 10.11.12.13 -HostName paula-laptop -ForceFQDNMatch

        BEGIN: Match PTR to hostname (force FQDN match)

        Invalid FQDN 'paula-laptop'; please provide a fully-qualified domain name (e.g., foo.bar.com) when using -ForceFQDNMatch

        IPAddress     : 10.11.12.13
        Hostname      : paula-laptop
        ReverseLookup : Error
        IsMatch       : Error
        DNSServers    : {10.60.157.250, 10.61.179.250, 10.62.179.250}
        Messages      : Invalid FQDN 'paula-laptop'

        END  : Match PTR to hostname (force FQDN match)


#>
[CmdletBinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Hostname to match (default is local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$HostName,
    
    [Parameter(
        Mandatory = $True,
        Position = 1,
        HelpMessage = "IPv4 address (default is local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$IPAddress,

    [Parameter(
        HelpMessage = "Require exact match for hostname's FQDN and PTR (default is hostname only)"
    )]
    [switch]$ForceFQDNMatch,

    [Parameter(
        HelpMessage = "Return full or Boolean output (default is full)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Full","Boolean")]
    [String]$OutputType = "Full",

    [Parameter(
        HelpMessage = "Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Error handling
    $ErrorActionPreference = "Stop"
    
    # How did we get here?
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    
    # We can't specify the DNS servers to use with the .NET methods; this is for reporting
    $WMINet = (Get-CIMInstance -ClassName win32_networkadapterconfiguration -filter "ipenabled = 'True'" -Verbose:$False -ErrorAction Stop)
    $DNSServers = $WMINet.DNSServerSearchOrder | Where-Object {$_ -notmatch ":"}

    # Not using default parameter values
    If (-not $CurrentParams.IPAddress) {
        $CurrentParams.IPAddress = $IPAddress = ($WMINet.IPAddress | Where-Object {$_ -notmatch ":"})
    }
    ElseIf (-not $CurrentParams.Hostname) {
        # https://boerlowie.wordpress.com/2010/12/31/get-the-fqdn-of-your-host-with-powershell/
        $objIPProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()
        $CurrentParams.Hostname = $Hostname = "{0}.{1}" -f $objIPProperties.HostName, $objIPProperties.DomainName
    }

    # Show our settings
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("DNSServers",$DNSServers)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

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

    #region Output object
    
    $InitialValue = "Error"
    $OutputTemplate = [pscustomobject] @{
        Hostname      = $InitialValue
        IPAddress     = $InitialValue
        ReverseLookup = $InitialValue
        IsMatch       = $InitialValue
        DNSServers    = $DNSServers
        Messages      = $InitialValue
    }
    $Select = $OutputTemplate.PSObject.Properties.Name
    If ($OutputType -eq "Boolean") {$Select = $OutputTemplate.IsMatch}

    #endregion Output object

    # Regex for valid hostnames (HT to https://www.regextester.com/98986)
    $FQDNRegex = "^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.){2,}([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9]){2,}$"

    # Console output
    $Activity = "Match PTR to hostname"
    If ($ForceFQDNMatch.IsPresent) {$Activity += " (enforce FQDN match)"}
    "BEGIN: $Activity" | Write-MessageInfo -FGColor Yellow -Title

}
Process {
    
    [switch]$Continue = $True

    # Make a copy of the output object
    $Output1 = $OutputTemplate.PSObject.Copy()
    $Output1.Hostname = $Hostname
    
    If ($ForceFQDNMatch.IsPresent -and ($Hostname -notmatch $Regex)) {
        $Msg = "Invalid FQDN '$Hostname'"
        $Output1.Messages = $Msg
        $Msg += "; please provide a fully-qualified domain name (e.g., foo.bar.com) when using -ForceFQDNMatch"
        $Msg | Write-MessageError 

        $Continue = $False
    }
    
    If ($Continue.IsPresent) {
        
        $Hostname = ($Hostname | Foreach-Object {$_.trim()})
        $Output1.Hostname = $Hostname
     
        # Make sure it's an ip address (not using parameter validation for better error handling)   
        

        Try {
            $Msg = "Resolve name in DNS"
            "[$Hostname] $Msg" | Write-MessageInfo -FGColor White
            [array]$IPAddress = [System.Net.Dns]::GetHostAddresses($Hostname).IPAddressToString | Where-Object {$_ -notmatch ":"}

            Foreach ($IP in $IPAddress) {
                
                # Make an 'inner' copy of the output object    
                $Output2 = $Output1.PSObject.Copy()
                $Output2.IPAddress = $IP
                $Msg = "Perform reverse lookup on IP address '$IP'"
                "[$Hostname] $Msg" | Write-MessageInfo -FGColor White

                Try {
                    $ErrorActionPreference = "SilentlyContinue"
                    $PTR = [System.Net.Dns]::GetHostEntry($IP).Hostname
                    $Output2.ReverseLookup = $PTR
                    $ErrorActionPreference = "Stop"

                    [switch]$IsMatch = $False
                    Switch ($ForceFQDNMatch) {
                        $True  {
                            If ($PTR -eq $Hostname) {$IsMatch = $True}
                        }
                        $False {
                            If ($PTR -match "^($Hostname\.)?") {$IsMatch = $True}
                        }
                    }
                    $Output2.IsMatch = $IsMatch
                    Switch ($IsMatch) {
                        $True {
                            $FGColor = "Green"
                            $Msg = "PTR/hostname match"
                        }
                        $False {
                            $FGColor = "Red"
                            $Msg = "PTR/hostname mismatch"
                        }
                    }
                    $Output2.Messages = $Msg
                    "[$Hostname] $Msg" | Write-MessageInfo -FGColor $FGColor
                }
                Catch {
                    $Msg = "PTR not found"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[$Hostname] $Msg" | Write-MessageError
                    $Output2.IsMatch = $False
                    $Output2.Messages = $Msg
                }
                    
                Write-Output ($Output2 | Select-Object $Select)
            }
        }
        Catch {
            $Msg = "Failed to resolve name"
            "[$Hostname] $Msg" | Write-MessageError
            $Output1.Messages = $Msg
            $Output1.IsMatch = $False
            Write-Output ($Output1 | Select-Object $Select)
        }
            <#
            Catch {
                $Msg = "Failed to resolve name"
                "[$Hostname] $Msg" | Write-MessageError
                $Output1.Messages = $Msg
                $Output1.IsMatch = $False
            }

            Try {
                $Msg = "Look up PTR record"
                "[$IPAddress] $Msg" | Write-MessageInfo -FGColor White
    
                $ErrorActionPreference = "SilentlyContinue"
                $PTR = [System.Net.Dns]::GetHostEntry($IPAddress).Hostname
                $Output.ReverseLookup = $PTR
                $ErrorActionPreference = "Stop"

                [switch]$IsMatch = $False
                Switch ($ForceFQDNMatch) {
                    $True  {
                        If ($PTR -eq $Hostname) {$IsMatch = $True}
                    }
                    $False {
                        If ($PTR -match "^($Hostname\.)?") {$IsMatch = $True}
                    }
                }
                $Output.IsMatch = $IsMatch
                Switch ($IsMatch) {
                    $True {
                        $FGColor = "Green"
                        $Msg = "PTR/hostname match"
                    }
                    $False {
                        $FGColor = "Red"
                        $Msg = "PTR/hostname mismatch"
                    }
                }
                $Output.Messages = $Msg
                "[$IPAddress] $Msg" | Write-MessageInfo -FGColor $FGColor
            }
            Catch {
                $Msg = "PTR not found"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$IPAddress] $Msg" | Write-MessageError
                $Output.IsMatch = $False
                $Output.Messages = $Msg
            }    
        }
        Catch {
            $Msg = "Invalid IPv4 syntax"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
            "[$IPAddress] $Msg" | Write-MessageError
            $Output.IsMatch = $False
            $Output.Messages = $Msg
        }

        #>
    }

}

End {

    "END  : $Activity" | Write-MessageInfo -FGColor Yellow -Title
}
} #end Compare-PKHostnameToPTR

