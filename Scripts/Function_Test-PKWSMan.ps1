#Requires -version 3
Function Test-PKWinRM {
<# 
.SYNOPSIS
    Test WinRM connectivity to a remote computer using various protocols

.DESCRIPTION
    Test WinRM connectivity to a remote computer using various protocols
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Test-PKWSMan.ps1
    Created : 2018-10-23
    Author  : Paula Kingsley
    Version : 02.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-10-23 - Created script
        v01.01.0000 - 2019-03-26 - Minor updates
        v02.00.0000 - 2019-10-07 - Renamed from Test-PKWSMan (created alias), overhauled, more updates

.PARAMETER ComputerName
    One or more computer names

.PARAMETER Credential
    Valid credentials on target (default is current user credentials)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Test-PKWinRM -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        Verbose        True                                     
        ComputerName   LAPTOP                          
        Credential     System.Management.Automation.PSCredential
        Authentication Negotiate                                
        TestPing       False                                    
        LookupDNS      False                                    
        BooleanOutput  False                                    
        Quiet          False                                    
        PipelineInput  False                                    
        ScriptName     Test-PKWinRM                             
        ScriptVersion  1.2.0                                    

        WARNING: [Prerequisites] Note that connections may fail unless fully-qualified domain names are provided; 
        try -LookupDNS if you want to attempt to get the FQDN from a hostname

        BEGIN  : Test WSMAN connectivity using Negotiate authentication

        [LAPTOP] Test WinRM connection

        ComputerName   : LAPTOP
        IsSuccess      : True
        PingResults    : -
        Username       : jbloggs
        Authentication : Negotiate
        Messages       : Successfully connected using WinRM and Negotiate authentication



.EXAMPLE
    PS C:\> $DCList | Test-PKWinRM -Credential $AdminCred -Authentication Kerberos -TestPing -Quiet -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        Credential     System.Management.Automation.PSCredential
        TestPing       True                                     
        Quiet          True                                     
        Verbose        True                                                                     
        ComputerName                                            
        Authentication Kerberos                                
        LookupDNS      False                                    
        BooleanOutput  False                                    
        PipelineInput  True                                     
        ScriptName     Test-PKWSMan                             
        ScriptVersion  1.2.0                                    

        WARNING: [Prerequisites] Note that connections may fail unless fully-qualified domain names are provided; 
        try -LookupDNS if you want to attempt to get the FQDN from a hostname

        VERBOSE: BEGIN : Test WSMAN connectivity using Kerberos authentication
        
        VERBOSE: [dc04.domain.local] Test ping
        VERBOSE: [dc04.domain.local] Ping successful
        VERBOSE: [dc04.domain.local] Test WinRM connection

        ComputerName   : dc04.domain.local
        IsSuccess      : True
        PingResults    : True
        Username       : DOMAIN\jbloggs-admin
        Authentication : Kerberos
        Messages       : Successfully connected using WinRM and Kerberos authentication

        VERBOSE: [dc03.domain.local] Test ping
        VERBOSE: [dc03.domain.local] Ping successful
        VERBOSE: [dc03.domain.local] Test WinRM connection
        ComputerName   : dc03.domain.local
        IsSuccess      : True
        PingResults    : True
        Username       : DOMAIN\jbloggs-admin
        Authentication : Kerberos
        Messages       : Successfully connected using WinRM and Kerberos authentication

        VERBOSE: [dc02.domain.local] Test ping
        VERBOSE: [dc02.domain.local] Ping successful
        VERBOSE: [dc02.domain.local] Test WinRM connection
        ComputerName   : dc02.domain.local
        IsSuccess      : False
        PingResults    : True
        Username       : DOMAIN\jbloggs-admin
        Authentication : Kerberos
        Messages       : The client cannot connect to the destination specified in the request. Verify that the service on the destination is running and is accepting requests. Consult the logs and documentation for the WS-Management service running on the 
                         destination, most commonly IIS or WinRM. If the destination is the WinRM service, run the following command on the destination to analyze and configure the WinRM service: "winrm quickconfig".

        VERBOSE: [dc01] Test ping
        VERBOSE: [dc01] Ping successful
        VERBOSE: [dc01] Test WinRM connection
        ComputerName   : dc01
        IsSuccess      : True
        PingResults    : True
        Username       : DOMAIN\jbloggs-admin
        Authentication : Kerberos
        Messages       : Successfully connected using WinRM and Kerberos authentication

        VERBOSE: END   : Test WSMAN connectivity using Kerberos authentication

.EXAMPLE
    PS: C:\> Test-PKWinRM -ComputerName sqldevbox -Authentication CredSSP -TestPing -Lookupdns -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        ComputerName   {sqldevbox}                        
        Authentication CredSSP                                  
        TestPing       True                                     
        LookupDNS      True                                     
        Verbose        True                                     
        Credential     System.Management.Automation.PSCredential
        BooleanOutput  False                                    
        Quiet          False                                    
        PipelineInput  False                                    
        ScriptName     Test-PKWinRM                             
        ScriptVersion  1.2.0                                    

        BEGIN  : Test WSMAN connectivity using CredSSP authentication

        [sqldevbox] Lookup hostname in DNS
        [sqldevbox.lab.internal.com] Resolved hostname to FQDN
        [sqldevbox.lab.internal.com] Test ping
        [sqldevbox.lab.internal.com] Ping successful
        [sqldevbox.lab.internal.com] Test WinRM connection
        [sqldevbox.lab.internal.com] WinRM test failed

        ComputerName   : sqldevbox.lab.internal.com
        IsSuccess      : False
        PingResults    : True
        Username       : scaruso
        Authentication : CredSSP
        Messages       : The WinRM client cannot process the request. Requests must include user name and password when CredSSP authentication 
                         mechanism is used. Add the user name and password or change the authentication mechanism and try the request again. 

        END  : Test WSMAN connectivity using CredSSP authentication

.EXAMPLE
    PS C:\> Test-PKWinRM -ComputerName sqldevbox -Authentication CredSSP -TestPing -Lookupdns -Quiet -BooleanOutput
        False

        
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
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

    [Parameter(
        HelpMessage="Valid credentials on target"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage="Authentication protocol: Basic, CredSSP, ClientCertificate, Digest, Kerberos,Negotiate, None (default Negotiate)"
    )]
    [ValidateSet("Basic","CredSSP","ClientCertificate","Digest","Kerberos","Negotiate","None")]
    [string] $Authentication = "Negotiate",

    [Parameter(
        HelpMessage = "Ping target"
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $TestPing,

    [Parameter(
        HelpMessage = "Attempt to look up hostname in DNS if FQDN not provided"
    )]
    [Switch] $LookupDNS,

    [Parameter(
        HelpMessage = "Return Boolean output only"
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $BooleanOutput,

    [Parameter(
        HelpMessage ="Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    $CurrentParams = $PSBoundParameters
    If (-not $PipelineInput -and (-not $ComputerName)) {
        $Computername = $CurrentParams.ComputerName = $Env:ComputerName
    }

    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    If (-not $LookupDNS.IsPresent) {
        $Msg = "Note that connections may fail unless fully-qualified domain names are provided; try -LookupDNS if you want to attempt to get the FQDN from a hostname"
        Write-Warning "[Prerequisites] $Msg"
    }

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
        Param([Parameter(ValueFromPipeline)]$Message)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("$Message")}
        Else {Write-Verbose "$Message"}
    }

    # Function to test WinRM connectivity
    Function Test-WinRM{
        Param($Computer)
        $Param_WSMAN = @{
            ComputerName   = $Computer
            Credential     = $Credential
            Authentication = $Authentication
            ErrorAction    = "Silentlycontinue"
            Verbose        = $False
            ErrorVariable  = "Failed"
        }
        If (-not $Ping) {$Ping = "-"}
        Try {
            $Test = Test-WSMan @Param_WSMAN
            If ($Test) {
                [pscustomobject]@{
                    ComputerName   = $Computer
                    IsSuccess      = $True
                    PingResults    = $Ping
                    Username       = $Username
                    Authentication = $Authentication
                    Messages       = "Successfully connected using WinRM and $Authentication authentication"
                }       
            }
            Else {
                [pscustomobject]@{
                    ComputerName   = $Computer
                    IsSuccess      = $False
                    PingResults    = $Ping
                    Username       = $Username
                    Authentication = $Authentication
                    Messages       = $([regex]:: match($Failed,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim())
                }
            }
        }
        Catch {
            [pscustomobject]@{
                    ComputerName   = $Computer
                    IsSuccess      = $False
                    PingResults    = $Ping
                    Username       = $Username
                    Authentication = $Authentication
                    Messages       = $_.Exception.Message
            }    
        }
    } #end Test-WinRM

    # Function to look up IP/Hostname
    Function Lookup-DNS($Target,[switch]$GetName) {
        $VerbosePreference = "SilentlyContinue"
        $ErrorActionPreference = "SilentlyContinue"
        Try {
            If ([bool]($Target -is [ipaddress])) {
                Resolve-DNSName -Name $Target -verbose:$False | Select -ExpandProperty NameHost
            }
            Elseif ($GetName.IsPresent) {
                Resolve-DNSName -Name $Target | Select -ExpandProperty Name
            }
            Else {
                Resolve-DNSName -Name $Target | Select -ExpandProperty IPAddress
            }
        }
        Catch {
            $_.Exception.Message
        }
    }

    # Function to test ping connectivity
    Function Test-Ping{
        Param($Target)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Target)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
        $Task.Dispose()
    }

    #endregion Functions

    #region Username

    If ($CurrentParams.Credential.Username) {
        $Username = $Credential.Username
    }
    Else {
        $Username = $Env:Username
    }
    
    #endregion Username

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Test-WSMan
    $Param_WSMAN = @{}
    $Param_WSMAN = @{
        Computer       = $Null
        Authentication = $Authentication
        BooleanOutput  = $BooleanOutput
        ErrorAction    = "Silentlycontinue"
        Verbose        = $False
    }

    # Splat for Write-Progress
    $Activity = "Test WSMAN connectivity using $Authentication authentication"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    #endregion Splats

    # Console output
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title


} #end begin

Process {
    
    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0
    
    Foreach ($Computer in $ComputerName) {
        
        $Current ++ 
        $Param_WP.PercentComplete = ($Current/$Total* 100)
        $Param_WP.Status = $Computer
        
        If ($Computer -is [string]) {
            $Computer = $Computer.Trim()
        }
        Elseif ($Computer -is [Microsoft.ActiveDirectory.Management.ADAccount]) {
            If ($Computer.DNSHostName) {
                $Computer = $Computer.DNSHostName
            }
            Else {
                $Computer = $Computer.Name
            }
        }
        
        If ((-not [bool]($Computer -is [ipaddress])) -and ($Computer -notmatch "\.")) {

            If ($LookupDNS.IsPresent) {
                $Msg = "Lookup hostname in DNS"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                Try {
                    If ($Target = Lookup-DNS -Target $Computer -GetName) {
                                    
                        $Computer = $Target
                        $Msg = "Resolved hostname to FQDN"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Green
                    }
                    Else {
                        $Msg = "Failed to resolve hostname to FQDN"
                        "[$Computer] $Msg" | Write-MessageError
                    }
                }
                Catch {}
            }
        } #end if not FQDN

        If ($TestPing.IsPresent) {
            $Msg = "Test ping"
            "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP

            If ($Ping = Test-Ping -Target $Computer) {
                $Msg = "Ping successful"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor Green
            }
            Else {
                $Msg = "Ping failure"
                "[$Computer] $Msg" | Write-MessageError
            }
        }


        $Msg = "Test WinRM connection"
        "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
        $Param_WP.CurrentOperation = $Msg
        Write-Progress @Param_WP

        $Param_WSMAN.Computer = $Computer
        $TestWinRM = Test-WinRM @Param_WSMAN
        
        Switch ($TestWinRM.IsSuccess) {
            $True   {
                $Msg = "WinRM test successful"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor Green
            }
            $False  {
                $Msg = "WinRM test failed"
                "[$Computer] $Msg" | Write-MessageError
            }
            Default {
                $Msg = "WinRM test error"
                "[$Computer] $Msg" | Write-MessageError
            }
        }
    
        If ($BooleanOutput.IsPresent) {
            Write-Output $TestWinRM.IsSuccess
        }
        Else {
            Write-Output $TestWinRM
        }        

    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    $Msg = "END  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title
    
}

} # end Test-PKWSMan


$Null = New-Alias -Name Test-PKWSMan -Value Test-PKWinRM -Confirm:$False -Force -Verbose:$False