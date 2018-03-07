#requires -Version 3
Function Test-PKDNSResolution {
<#
.SYNOPSIS
    Resolves a name to an IP address, or an IP address to a hostname, using .NET system.net.dns methods

.DESCRIPTION
    Resolves a name to an IP address, or an IP address to a hostname, using .NET system.net.dns methods
    Uses [System.Net.Dns]::GetHostEntryAsync() to speed performance
    Uses locally configured nameservers (cannot be specified)
    Accepts pipeline input
    Returns a PSObject

.NOTES
    Name    : Function_Test-PKDNSResolution.ps1
    Author  : Paula Kingsley
    Created : 2018-01-26
    Version : 01.00.0000    
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2018-01-26 - Created script


.PARAMETER InputObj
    Name, FQDN, or IP address to look up

.PARAMETER SuppressConsoleOutput
    Suppress all non-verbose/non-error console output

.PARAMETER ReturnTrueOnly
    Return results only for entries that successfully resolve

.EXAMPLE
    PS C:\> Test-PKDNSResolution -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value               
        ---                   -----               
        Verbose               True                
        InputObj                                  
        ReturnTrueOnly        False               
        SuppressConsoleOutput False               
        PipelineInput         True                
        ScriptName            Test-PKDNSResolution
        ScriptVersion         1.0.0               

        Action: Test DNS name/IP resolution
        VERBOSE: WORKSTATION14

        Input          IsResolved Output                                                
        -----          ---------- ------                                                
        WORKSTATION14  True       {10.60.197.240, 192.168.56.1, 192.168.232.1, 192.16...

.EXAMPLE
     PS C:\> $Arr | Test-PKDNSResolution
        
        Action: Test DNS name/IP resolution

        Input             IsResolved Output                  Errors           
        -----             ---------- ------                  ------           
        4.2.2.1           True       a.resolvers.level3.net                  
        foo               False                              Resolution failed
        !@$%^^            False                              Invalid syntax   
        dc1.domain.local  True       10.14.11.32

.EXAMPLE
    PS C:\> $Arr | Test-PKDNSResolution -ReturnTrueOnly

        Action: Test DNS name/IP resolution (return only valid results)

        Input             IsResolved Output                  Errors           
        -----             ---------- ------                  ------           
        4.2.2.1           True       a.resolvers.level3.net                  
        dc1.domain.local  True       10.14.11.32


.EXAMPLE
    PS C:\> Test-PKDNSResolution -InputObj host.domain.local -SuppressConsoleOutput

        Input             IsResolved Output       Errors
        -----             ---------- ------       ------
        host.domain.local True       10.11.178.23       

#>
[CmdletBinding()]
Param(
    [Parameter(
        Position = 0,
        Mandatory = $False,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Hostname or IP"
    )]
    [Alias("Name","IP","DNSHostName","FQDN")]
    [ValidateNotNullOrEmpty()]
    [object[]]$InputObj,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Return resugit status
        lts only for successfully resolved entries"
    )]
    [switch]$ReturnTrueOnly,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Suppress all non-verbose/non-error console output"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$SuppressConsoleOutput
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Detect pipeline input and save parametersetname
    $Source = $PSCmdlet.ParameterSetName
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("InputObj")) -and ((-not $InputObj)) # -or (-not $InputObj -eq $Env:ComputerName))

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # If we didn't supply anything
    If (-not $InputObj) {$InputObj = $Env:ComputerName}

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    #Splat for Write-Progress
    $Activity = "Test DNS name/IP resolution"
    $Param_WP = @{
        Activity         = $Activity
        Status           = "Working"
        CurrentOperation = $Null
        PercentComplete  = $Null
    }
    
    #Output 
    [array]$Results = @()
    $InvalidNameRegex = '[\\%\/:"*?<>@)(~#&$|]+'
    $ValidIPRegex = "\b(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}\b"
    
    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Action = $Activity
    If ($ReturnTrueOnly.IsPresent) {$Action += " (return only valid results)"}
    $Msg = "Action: $Action"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}

Process {
    
    $Lookup = @()
    $OutputObj = @()

    # Remove dupes
    $InputObj = $InputObj | Select-Object -Unique

    $Total = $InputObj.Count
    $Current = 0

    Try {
        [array]$Lookup = $InputObj | Select-Object @{N="Input";E={$_}},@{N="Test";E={[System.Net.Dns]::GetHostEntryAsync($_)}}

        $Lookup | Foreach-Object  {
            
            $Current ++

            $Curr = $_
            $Msg = $Curr.Input
            Write-Verbose $Msg

            $Param_WP.CurrentOperation = $Msg
            $Param_WP.PercentComplete = ($Current/$Total * 100)
            $Param_WP.Status = "$([math]::Round($Current/$Total * 100))%"
            Write-Progress @Param_WP
            
            If ($Curr.Test.Result) {
                [switch]$IsResolved = $True
                $Output = Switch -Regex ($Curr.Input) {
                    $ValidIPRegex {$Curr.Test.Result.Hostname}
                    Default {$Curr.Test.Result.AddressList}#.IPAddressToString}
                }
                $ErrMsg = $Null
            }
            Else {
                [switch]$IsResolved = $False
                $Output = $Null                    
                If ($Curr.Input -match $InvalidNameRegex) {$ErrMsg = "Invalid syntax"}
                Else {$ErrMsg = "Resolution failed"}
            
            }

            $Results += New-Object PSObject -Property ([ordered]@{
                Input      = $Curr.Input
                IsResolved = $IsResolved
                Output     = $Output
                Errors     = $errMsg
            })
    
        } #end for each
    
    }
    Catch {
        $Msg = "Name resolution test failure"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
        $Host.UI.WriteErrorLine($Msg)
        $Results += New-Object PSObject -Property ([ordered]@{
            Input      = $InputObj
            IsResolved = $False
            Output     = "Error"
            Errors     = $Msg
        })
    }

}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    
    # Get rid of empty lines, just in case
    $Results = ($Results | Where-Object {$_.Input})

    If ($ReturnTrueOnly.IsPresent) {Write-Output $Results | Where-Object {$_.IsResolved}}
    Else {Write-Output $Results}

}
} #end Test-PKDNSResolution.ps1



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