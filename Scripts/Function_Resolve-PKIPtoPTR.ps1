#Requires -version 3
Function Resolve-PKIPtoPTR {
<# 
.SYNOPSIS
    Performs a reverse DNSlookup on one or more IPv4 addresses, against one or more DNS server names or IP addresses

.DESCRIPTION
    Performs a reverse DNSlookup on one or more IPv4 addresses, against one or more DNS server names or IP addresses
    Optionally runs as a job
    Defaults to locally configured IP addresses if not specified
    Defaults to locally configured DNS server IPs if  not specified
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Resolve-PKIPtoPTR.ps1
    Created : 2019-12-16
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-12-16 - Created script

.PARAMETER IPAddress
    One or more IPv4 addresses (default is local computer)

.PARAMETER IPAddress
    One or more DNS servers (default is all currently configured)

.PARAMETER NoRecursion
    Perform non-recursive lookup only

.PARAMETER AsJob
    Run as a job

.PARAMETER JobName
    Prefix for job (default is PTR)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\>  Resolve-PKIPtoPTR -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                        
        ---              -----                                        
        Verbose          True                                         
        Server           {172.30.28.30, 10.64.8.197}
        IPAddress        {172.30.2.59}                              
        NoRecursion      False                                        
        Quiet            False                                        
        PipelineInput    False                                                               
        ScriptName       Resolve-PKIPtoPTR                            
        ScriptVersion    1.0.0                                        

        BEGIN: Perform reverse lookup

        VERBOSE: [172.30.2.59] Perform reverse lookup using DNS server 172.30.28.30

        IPAddress : 172.30.2.59
        Name      : 59.2.30.172.in-addr.arpa
        Resolved  : True
        NameHost  : WORKSTATION7.domain.local
        Server    : 172.30.28.30
        Status    : Successfully resolved PTR

        VERBOSE: [172.30.2.59] Perform reverse lookup using DNS server 10.64.8.197
        IPAddress : 172.30.2.59
        Name      : 59.2.30.172.in-addr.arpa
        Resolved  : True
        NameHost  : WORKSTATION7.domain.local
        Server    : 10.64.8.197
        Status    : Successfully resolved PTR

        END  : Perform reverse lookup

.EXAMPLE
    PS C:\> Resolve-PKIPtoPTR -IPAddress 10.30.22.242,google.com -Server 192.168.30.30,4.2.2.1 -Quiet

        IPAddress : 10.30.22.242
        Name      : 242.22.30.10.in-addr.arpa
        Resolved  : True
        NameHost  : sql-test-5.corp.net
        Server    : 192.168.30.30
        Status    : Successfully resolved PTR

        IPAddress : 10.30.22.242
        Name      : Error
        Resolved  : False
        NameHost  : Error
        Server    : 4.2.2.1
        Status    : 242.22.30.10.in-addr.arpa : DNS name does not exist

        IPAddress : google.com
        Name      : Error
        Resolved  : False
        NameHost  : Error
        Server    : 4.2.2.1
        Status    : Input is not a valid IP address
        
#> 

[CmdletBinding(
    DefaultParameterSetName = "__DefaultParameterSet",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more IPv4 addresses (default is local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [String[]] $IPAddress = @(@(Get-CIMInstance -Class Win32_NetworkAdapterConfiguration -Verbose:$False -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress) -like "*.*"),

    [Parameter(
        HelpMessage="One or more DNS server names or IPs"
    )]
    [ValidateNotNullOrEmpty()]
    [String[]] $Server,

    [Parameter(
        HelpMessage = "Perform non-recursive lookup only"
    )]
    [Switch] $NoRecursion,

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Run as job"
    )]
    [Switch] $AsJob,

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Prefix for job name (default is PTR)"
    )]
    [string]$JobPrefx = "PTR",

    [Parameter(
        HelpMessage="Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    $CurrentParams = $PSBoundParameters
    If (-not $CurrentParams.Server) { 
        $CurrentParams.Server = $Server = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE | Select-Object -Expand DNSServerSearchOrder)
    }
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    #region Prerequisite check
    Try {
        $Null = Get-Command -Name Resolve-DNSName -ErrorAction Stop
    }
    Catch { # Not setting Requires for version 4 because other functions in module don't require it
        $Msg = "This function requires PowerShell 4 on Windows 8/2012 at minimum; cmdlet Resolve-DNSName not found"
        Throw $Msg
    }

    #endregion Prerequisite check

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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    #endregion Functions

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Write-Progress
    $Activity = "Perform reverse lookup"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    # Splat for Resolve-DNSName
    $Param_RD = @{}
    $Param_RD = @{
        Name             = $Null
        Type             = "PTR"
        Server           = $Null
        NoRecursion      = $NoRecursion
        ErrorAction      = "Stop"
        Verbose          = $False
    }
    
    #endregion Splats

    # Console output
    "BEGIN: $Activity" | Write-MessageInfo -FGColor Yellow -Title
    


} #end begin

Process {
    
    $Total = $IPAddress.Count
    $Current = 0

    Foreach ($IP in ($IPAddress | Select-Object -Unique)) {
        
        Try {$IP = $IP.Trim()}
        CatcH {}

        If ($IP -as [ipaddress]) {
            
            Foreach ($S in ($Server | Select-Object -Unique)) {
                
                $Current ++
                $Param_WP.Status = $IP 
                $Param_WP.CurrentOperation = "DNS server $S" 
                $Param_WP.PercentComplete = ($Current/$Total)

                Write-Verbose "[$IP] Perform reverse lookup using DNS server $S"

                Try {
                    Write-Progress @Param_WP

                    $Param_RD.Name = $IP.Trim()
                    $Param_RD.Server = $S
                    
                    #$DNS = Resolve-DnsName @Param_RD

                    If ($AsJob.IsPresent) {
                    
                    
                    }

                    Start-Job -ArgumentList $Param_RD,$IP,$S -ScriptBlock {
                    Param($Param_RD,$IP,$S)
                    Resolve-DnsName @Using:Param_RD | 
                        Where-Object Section -eq Answer | 
                            Sort-Object NameHost | Group-Object Name |
                                Select-Object @{N='IPAddress';E={$Using:IP}},
                                Name,
                                @{N='Resolved';E={$True}},
                                @{N="NameHost";E={$_.Group.NameHost}},
                                @{N='Server';E={$Using:S}},
                                @{N='Status';E={'Successfully resolved PTR'}}
                    }
              
                }
                Catch{
                    $ErrorDetails = $_.Exception.Message
                    $ErrorDetails | Write-MessageError 
                    [pscustomobject]@{
                        IPAddress = $IP
                        Name      = "Error"
                        Resolved  = $False
                        NameHost  = "Error"
                        Server    = $S
                        Status    = $ErrorDetails #$Error[0].Exception.Message
                    }
                }
            } #end for each server
        
        } #end if valid IP format
        Else {
            [pscustomobject]@{
                IPAddress = $IP
                Name      = "Error"
                Resolved  = $False
                NameHost  = "Error"
                Server    = $S
                Status    = "Input is not a valid IP address"
            }
        }
    } # end for each IP

}
End {
    
    $Null = Write-Progress -Activity * -Completed
    "END  : $Activity" | Write-MessageInfo -FGColor Yellow -Title
}

} # end Resolve-PKIPtoPTR

