#Requires -version 3
Function Get-PKIPConfig {
<# 
.SYNOPSIS
    Gets IPv4 information for one or more Windows computers, interactively or as a PSJob

.DESCRIPTION
    Gets IPv4 information for one or more Windows computers, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob
    NOTE: Windows 10 may return erroneous date/time for DHCP leases!

.NOTES        
    Name    : Function_Get-PKIPConfig.ps1.ps1
    Created : 2019-11-20
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-11-20 - Created script

.PARAMETER ComputerName
    One or more computer names

.PARAMETER AllAdapters
    Return data for all adapters (default is active/IPEnabled only)

.PARAMETER Credential
    Valid credentials on target (default is current user)

.PARAMETER Authentication
    WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)

.PARAMETER AsJob
    Invoke command as a PSjob

.PARAMETER JobPrefix
    Prefix for job name (default is 'IP')

.PARAMETER ConnectionTest
    Run WinRM or ping connectivity test prior to Invoke-Command, or no test (default is WinRM)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> Get-PKIPConfig -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        ComputerName   {DEVBOX}                        
        Credential     System.Management.Automation.PSCredential
        Verbose        True                                     
        AllAdapters    False                                    
        Authentication Negotiate                                
        AsJob          False                                    
        JobPrefix      IP                                       
        ConnectionTest WinRM                                    
        Quiet          False                                    
        ScriptName     Get-PKIPConfig                 
        ScriptVersion  2.0.0                                    
        PipelineInput  False                                    


        BEGIN: Invoke scriptblock to retrieve TCP/IP configuration details

        [DEVBOX] Invoke command

        ComputerName                 : DEVBOX
        OperatingSystem              : Microsoft Windows Server 2012 R2 Datacenter
        Domain                       : domain.local
        Description                  : vmxnet3 Ethernet Adapter
        NetConnectionID              : Ethernet0
        Index                        : 13
        MACAddress                   : 00:50:56:96:67:81
        Speed                        : 10000000000
        IPAddress                    : 192.168.57.49
        SubnetMask                   : 255.255.255.0
        DefaultGateway               : 192.168.57.1
        DNSHostName                  : devbox
        DNSDomain                    : domain.local
        ConnectionDNSSuffix          : lab.domain.local
        PrimaryDNS                   : 192.168.25.6
        DNSServerSearchOrder         : {192.168.25.6, 192.168.25.7}
        DNSDomainSuffixSearchOrder   : {domain.local, lab.domain.local}
        DomainDNSRegistrationEnabled : False
        FullDNSRegistrationEnabled   : True
        DHCPEnabled                  : True
        DHCPServer                   : 192.168.25.6
        DHCPLeaseObtained            : 10/12/2019 9:48:43 PM
        DHCPLeaseExpires             : 10/19/2019 9:48:43 PM
        DHCPLeaseTimeToLive          : 4.12:31:40.2734149
        Messages                     : 

        END  : Invoke scriptblock to retrieve TCP/IP configuration details

.EXAMPLE
    PS C:\> $Arr | Get-PKIPConfig -Credential $Credential -Authentication Kerberos -AsJob -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        AllAdapters    True                                     
        Credential     System.Management.Automation.PSCredential
        Authentication Kerberos                                 
        AsJob          True                                     
        Verbose        True                                     
        ComputerName                                            
        JobPrefix      IP                                       
        ConnectionTest WinRM                                    
        Quiet          False                                    
        ScriptName     Get-PKIPConfig                 
        ScriptVersion  2.0.0                                    
        PipelineInput  True                                     


        BEGIN: Invoke scriptblock to retrieve TCP/IP configuration details as remote PSJob

        [APPSERVER-2.domain.local] Test WinRM connection using Kerberos authentication
        [APPSERVER-2.domain.local] WinRM connection successful
        [APPSERVER-2.domain.local] Invoke command as PSJob
        [appserver-3.domain.local] Test WinRM connection using Kerberos authentication
        [appserver-3.domain.local] WinRM connection successful
        [appserver-3.domain.local] Invoke command as PSJob
        [SARAH-DEMO] Test WinRM connection using Kerberos authentication
        [SARAH-DEMO] WinRM failure
        [labmanager.domain.local] Test WinRM connection using Kerberos authentication
        [labmanager.domain.local] WinRM connection successful
        [labmanager.domain.local] Invoke command as PSJob
        [sqltest.lab.domain.local] Test WinRM connection using Kerberos authentication
        [sqltest.lab.domain.local] WinRM connection successful
        [sqltest.lab.domain.local] Invoke command as PSJob
        [sqltest2.lab.domain.local] Test WinRM connection using Kerberos authentication
        [sqltest2.lab.domain.local] WinRM connection successful
        [sqltest2.lab.domain.local] Invoke command as PSJob
        
        5 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name              PSJobTypeName   State         HasMoreData     Location           Command                  
        --     ----              -------------   -----         -----------     --------           -------                  
        26     IP_appserver-2.do RemoteJob       Completed     True            appserver-2.dom..  ...                      
        28     IP_appserver-3.do RemoteJob       Completed     True            appserver-3.dom... ...                      
        30     IP_labmanager.lab RemoteJob       Completed     True            labmanager.lab.... ...                      
        32     IP_sqltest.lab.do RemoteJob       Running       True            sqltest.lab.dom... ...                      
        34     IP_sqltest2.lab.d RemoteJob       Running       True            sqltest2.lab.do... ...                      

        END  : Invoke scriptblock to retrieve TCP/IP configuration details as remote PSJob

        PS C:\> Get-Job ip* | Receive-Job | Format-Table -Autosize

        ComputerName    OperatingSystem                             Domain            Description                 NetConnectionID Index MACAddress              Speed IPAddress     SubnetMask   
        ------------    ---------------                             ------            -----------                 --------------- ----- ----------              ----- ---------     ----------   
        APPSERVER-2     Microsoft Windows Server 2008 R2 Standard   domain.local      vmxnet3 Ethernet Adapter    eth0                7 00:50:56:96:7E:BB 10000000000 10.32.8.184   255.255.254.0
        APPSERVER-3     Microsoft Windows Server 2012 R2 Datacenter domain.local      vmxnet3 Ethernet Adapter #2 Ethernet1          11 00:50:56:96:2C:CC 10000000000 10.32.8.125   255.255.254.0
        LABMANAGER      Microsoft Windows Server 2012 R2 Datacenter lab.domain.local  vmxnet3 Ethernet Adapter    Ethernet0          10 00:50:56:96:7A:67 10000000000 10.32.8.235   255.255.254.0
        SQLTEST         Microsoft Windows Server 2012 R2 Datacenter lab.domain.local  vmxnet3 Ethernet Adapter    Ethernet0          13 00:50:56:96:67:81 10000000000 192.168.57.49 255.255.254.0
        SQLTEST2        Microsoft Windows Server 2012 R2 Datacenter lab.domain.local  vmxnet3 Ethernet Adapter    Ethernet0          10 00:50:56:96:33:AE 10000000000 10.32.8.158   255.255.254.0
    
.EXAMPLE
    PS C:\> Get-PKIPConfig -AllAdapters -Quiet -ConnectionTest None -Confirm:$False

        ComputerName                 : LAPTOP
        OperatingSystem              : Microsoft Windows 10 Pro
        Domain                       : domain.local
        Description                  : Microsoft Kernel Debug Network Adapter
        NetConnectionID              : 
        Index                        : 0
        MACAddress                   : 
        Speed                        : 
        IPAddress                    : 
        SubnetMask                   : 
        DefaultGateway               : 
        DNSHostName                  : 
        DNSDomain                    : 
        ConnectionDNSSuffix          : -
        PrimaryDNSServer             : 
        DNSServerSearchOrder         : 
        DNSDomainSuffixSearchOrder   : 
        DomainDNSRegistrationEnabled : 
        FullDNSRegistrationEnabled   : 
        DHCPEnabled                  : True
        DHCPServer                   : 
        DHCPLeaseObtained            : 
        DHCPLeaseExpires             : 
        DHCPLeaseTimeToLive          : 
        Messages                     : 

        ComputerName                 : LAPTOP
        OperatingSystem              : Microsoft Windows 10 Pro
        Domain                       : domain.local
        Description                  : Surface Ethernet Adapter
        NetConnectionID              : Ethernet 3
        Index                        : 7
        MACAddress                   : B8:31:B5:39:64:AC
        Speed                        : 1000000000
        IPAddress                    : 192.168.21.47
        SubnetMask                   : 255.255.255.0
        DefaultGateway               : 192.168.21.1
        DNSHostName                  : LAPTOP
        DNSDomain                    : domain.local
        ConnectionDNSSuffix          : domain.local
        PrimaryDNSServer             : 192.168.157.32
        DNSServerSearchOrder         : {192.168.157.32, 192.168.152.25}
        DNSDomainSuffixSearchOrder   : {parentcorp.lan, domain.local, datacenter.domain.loc...}
        DomainDNSRegistrationEnabled : False
        FullDNSRegistrationEnabled   : True
        DHCPEnabled                  : True
        DHCPServer                   : 192.168.157.32
        DHCPLeaseObtained            : 6/21/1933 11:32:52 PM
        DHCPLeaseExpires             : 6/26/1933 11:32:52 PM
        DHCPLeaseTimeToLive          : -31521.10:06:39.1912603
        Messages                     : WARNING: DHCP lease date/times returned via WMI may be incorrect for computers running Windows 10

        <snip>
        
          
#> 

[CmdletBinding(
    DefaultParameterSetName = "__DefaultParameterSet",
    SupportsShouldProcess = $True,
    ConfirmImpact = "Low"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

    [Parameter(
        HelpMessage = "Return data for all adapters (default is active/IPEnabled only)"
    )]
    [switch]$AllAdapters,
    
    [Parameter(
        HelpMessage = "Valid administrative credentials on target"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)"
    )]
    [ValidateSet('Kerberos','Basic','Negotiate','Default','CredSSP')]
    [string]$Authentication = "Negotiate",

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Run Invoke-Command scriptblock as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Prefix for job name (default is 'IP')"
    )]
    [String] $JobPrefix = "IP",

    [Parameter(
        HelpMessage = "Test to run prior to Invoke-Command - WinRM, ping, or none)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

    [Parameter(
        HelpMessage = "Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName
    }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
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

    # Function to write an error message (as a string with no stacktrace info)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    #endregion Functions
    
    #region Scriptblock for Invoke-Command
    
    $ScriptBlock = {
        Param($AllAdapters)
        # https://serverfault.com/questions/544116/find-computers-in-domain-with-static-ips-powershell
        # Not using [PSCustomObject] or Get-CIMInstance for backwards compatibility
        $ErrorActionPreference = "Stop"
        $InitialValue = "Error"
        
        $Param_GWMIComputer = @{
            ComputerName = "."
            Class        = "Win32_ComputerSystem"
            ErrorAction  = "Stop"
        }

        $Param_GWMIOS = @{
            ComputerName = "."
            Class        = "Win32_OperatingSystem"
            ErrorAction  = "Stop"
        }

        $Param_GWMINetAdapterConfig = @{
            ComputerName = "."
            Class        = "Win32_NetworkAdapterConfiguration"
            ErrorAction  = "Stop"
        }
        If (-not $AllAdapters.IsPresent) {
            $Param_GWMINetAdapterConfig.Add("Filter","IPEnabled='TRUE'")
        }

        $Param_GWMIAdapter = @{
            ComputerName = "."
            Class        = "Win32_NetworkAdapter"
            ErrorAction  = "Stop"
        }

        #region Select

        $Select = @{Name="ComputerName";Expression = {$ComputerObj.Name}},
            @{Name="OperatingSystem";Expression={$OS}},
            @{Name="Domain";Expression={$ComputerObj.Domain}},
            "Description",
            @{Name="NetConnectionID";Expression={(Get-WMIObject -Filter "Index=$($_.Index)" @Param_GWMIAdapter).NetConnectionID}},
            "Index",
            "MACAddress",
            @{Name="Speed";Expression={(Get-WmiObject -query "associators of {Win32_NetworkAdapterConfiguration.Index=$($_.index)}" -ErrorAction Stop).Speed}},
            @{Name="IPAddress";Expression={$_.IPAddress[0]}},
            @{Name="SubnetMask";Expression={$_.IPSubnet[0]}},
            @{Name="DefaultGateway";Expression={$_.DefaultIPGateway[0]}},
            "DNSHostName",
            "DNSDomain",
            @{Name="ConnectionDNSSuffix";Expression={
                If ($_.DNSDomain) {$_.DNSDomain}
                Else {"-"}
            }},
            @{Name="PrimaryDNSServer";Expression={$_.DNSServerSearchOrder[0]}},
            "DNSServerSearchOrder",
            "DNSDomainSuffixSearchOrder",
            "DomainDNSRegistrationEnabled",
            "FullDNSRegistrationEnabled",
            "DHCPEnabled",
            @{Name="DHCPServer";Expression = {
                If ($_.DHCPEnabled) {$_.DHCPServer}
                Else {"-"}
            }},
            @{Name="DHCPLeaseObtained";Expression={
                If ($_.DHCPEnabled) {$_.ConvertToDateTime($_.DHCPLeaseObtained)}
                Else {"-"}
            }},
            @{Name="DHCPLeaseExpires";Expression={
                If ($_.DHCPEnabled) {$_.ConvertToDateTime($_.DHCPLeaseExpires)}
                Else {"-"}
            }},
            @{Name="DHCPLeaseTimeToLive";Expression={ 
                If ($_.DHCPEnabled) {$_.ConvertToDateTime($_.DHCPLeaseExpires) - (Get-Date)}
                Else {"-"}
            }},
            @{Name="Messages";Expression={
                If ($OS -match "Windows 10" -and $_.DHCPEnabled -and $_.DHCPLeaseObtained) {
                    "WARNING: DHCP lease date/times returned via WMI may be incorrect for computers running Windows 10"
                }
                Else {$Null}
            }}

        #endregion Select

        Try {
            $ComputerObj = Get-WMIObject @Param_GWMIComputer | Select Name,Domain
            $OS = (Get-WMIObject @Param_GWMIOS).Caption
            $Output = Get-WmiObject @Param_GWMINetAdapterConfig | Select-Object $Select
        }
        Catch {
            $Msg = "Operation failed"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
            $Output = "" | Select-Object $Select
            $Output.Messages = $Msg
            $Output.PSobject.Properties | Foreach-Object {If (-not $_.Value -or $_.Value -eq "-") {$_.Value = "Error"}}
        }

        Write-Output $Output 

    } #end scriptblock

    #endregion Scriptblock for Invoke-Command

    #region Functions

    Function Test-WinRM{
        Param($Computer)
        $Param_WSMAN = @{
            ComputerName   = $Computer
            Credential     = $Credential
            Authentication = $Authentication
            ErrorAction    = "Silentlycontinue"
            Verbose        = $False
        }
        Try {
            If (Test-WSMan @Param_WSMAN) {$True}
            Else {$False}
        }
        Catch {$False}
    }

    Function Test-Ping{
        Param($Computer)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
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
    $Activity = "Invoke scriptblock to retrieve TCP/IP configuration details"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as PSJob"
    }
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    # Parameters for Invoke-Command
    $ConfirmMsg = $Activity
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName   = $Null
        Authentication = $Authentication
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $AllAdapters
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.Add("AsJob",$True)
        $Param_IC.Add("JobName",$Null)
    }

    # Parameters for Invoke-Command (local computer)
    $ConfirmMsg = $Activity
    $Param_IC_Local = @{}
    $Param_IC_Local = @{
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $AllAdapters
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC_Local.Add("AsJob",$True)
        $Param_IC_Local.Add("JobName",$Null)
    }
    
    #endregion Splats

    # Console output
    $Msg = "BEGIN: $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

    

} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
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

        [bool]$IsLocal = $False
        If ($Computer -in ($env:COMPUTERNAME,"localhost","127.0.0.1")) {
            $IsLocal = $True
        }
        
        $Current ++ 
        $Param_WP.PercentComplete = ($Current/$Total* 100)
        $Param_WP.Status = $Computer

        [switch]$Continue = $False

        Switch ($ConnectionTest) {
            Default {$Continue = $True}
            Ping {
                If (-not $IsLocal) {
                    $Msg = "Ping computer"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                        If ($Null = Test-Ping -Computer $Computer) {
                            $Continue = $True
                            $Msg = "Ping successful"
                            "[$Computer] $Msg" | Write-MessageInfo -FGColor Green
                        }
                        Else {
                            $Msg = "Ping failure"
                            "[$Computer] $Msg" | Write-MessageError
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageError
                    }
                }
                Else {$Continue = $True}
            }
            WinRM {
                If (-not $IsLocal) {
                    $Msg = "Test WinRM connection using $Authentication authentication"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White

                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                        If ($Null = Test-WinRM -Computer $Computer) {
                            $Continue = $True
                            $Msg = "WinRM connection successful"
                            "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                        }
                        Else {
                            $Msg = "WinRM failure"
                            "[$Computer] $Msg" | Write-MessageError
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageError
                    }
                }
                Else {$Continue = $True}
            }        
        }

        If ($Continue.IsPresent) {
            
            $ConfirmMsg = "`n`n`t$Activity`n`n"
            If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
                Try {
                    $Msg = "Invoke command"
                    If ($AsJob.IsPresent) {$Msg += " as PSJob"}
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    If ($IsLocal) {
                        If ($AsJob.IsPresent) {
                            $Job = $Null
                            $Param_IC_Local.JobName = "$JobPrefix`_$Computer"
                            $Job = Invoke-Command @Param_IC_Local 
                            $Jobs += $Job
                        }
                        Else {
                            Invoke-Command @Param_IC_Local | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID
                        } 
                    }
                    Else {
                        $Param_IC.ComputerName = $Computer
                        If ($AsJob.IsPresent) {
                            $Job = $Null
                            $Param_IC.JobName = "$JobPrefix`_$Computer"
                            $Job = Invoke-Command @Param_IC 
                            $Jobs += $Job
                        }
                        Else {
                            Invoke-Command @Param_IC | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID
                        }
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[$Computer] $Msg" | Write-MessageError
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                "[$Computer] $Msg" | Write-MessageError
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {

        If ($Jobs.Count -gt 0) {
            $Msg = "`n$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            "$Msg" | Write-MessageInfo -FGColor Green
            $Jobs | Get-Job
            
        }
        Else {
            $Msg = "No jobs created"
            "[$Computer] $Msg" | Write-MessageError
        }
    } #end if AsJob

    $Msg = "END  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}

} # end Get-PKIPConfig


New-Alias -Name Get-GNOpsWinIPConfig -Value Get-PKIPConfig -Force -Confirm:$False
New-Alias -Name Get-GNOpsIPConfig -Value Get-PKIPConfig -Force -Confirm:$False