#requires -version 3
Function Get-PKWindowsEvent {
<#
.SYNOPSIS
    Uses Get-WinEvent to return events from Windows event logs on one or more computers

.DESCRIPTION
    Uses Get-WinEvent to return events from Windows event logs on one or more computers
    Runs connectivity tests on remote computers (Ping, RPC connection on TCP port 135, or none)
    Defaults to current users' credentials; ignores credentials on local computer
    Uses -FilterHashtable to filter on start/end date, data string, level, etc.
    Created as a wrapper because building the hashtable manually is often annoying!
    Accepts pipeline input
    Returns a PSObject

.NOTES        
    Name    : Function_Get-PKWindowsEvent.ps1
    Created : 2022-02-22
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2022-02-22 - Created script
        
.PARAMETER ComputerName
    One or more Windows computer names or objects (default is local computer)

.PARAMETER LogName
    One or more event log names (default is Application, Security, and System; you can select * for all, but this will get extremely chatty!)

.PARAMETER EventType
    One or more severity levels for events: Critical, Error, Warning, Information, Warning (default is Critical, Error) 

.PARAMETER EventID
    One or more EventIDs (default is all matching event IDs)

.PARAMETER Data
    Text to search for in events (don't include asterisks; defaults to wildcard search)

.PARAMETER StartDate
    Earliest date for filtering matching events (default is past 5 days)

.PARAMETER EndDate
    Latest date for filtering matching events (default is now)

.PARAMETER MaxEvents
    Number of matching events to return (default is all)
    
.PARAMETER Credential
    Valid credentials on computer (ignored if local; note that authentication failure will still display a generic RPC error)

.PARAMETER ConnectionTest
    Connectivity test to run before attempting Get-WinEvent: Ping, RPC test on TCP 135, or None (default is Ping)

.EXAMPLE
    PS C:\> Get-PKWindowsEventLog -Verbose
    Returns Critical and Error messages from the Application & System event logs, within the past 5 days, on the local computer.

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        Verbose        True                                     
        ComputerName   LAPTOP                          
        LogName        {Application, Security, System}          
        EventType      {Critical, Error}                        
        EventID        0                                        
        Data                                                    
        StartDate      2022-03-17 3:17:47 PM                    
        EndDate        2022-03-22 3:17:48 PM                    
        MaxEvents      0                                        
        Credential     System.Management.Automation.PSCredential
        ConnectionTest Ping                                     
        ScriptName     Get-PKWindowsEventLog                    
        ScriptVersion  1.0.0                                    
        PipelineInput  False                                    

        VERBOSE: Running Get-WinEvent on remote computers requires valid credentials and an enabled inbound firewall rule for 'Remote Event Log Management.'
        (Unfortunately, remote authentication failures generate only generic 'RPC connection' failure messages.)
        You can enable the firewall rules via Group Policy, or via PowerShell ...it's safest to set this for the Domain profile only.

	        PS C:\> Get-NetFirewallRule -Name 'RemoteEventLogSvc-In-TCP' | Set-NetFirewallRule -Enabled True -Profile Domain -Direction Inbound -PassThru

        VERBOSE: [BEGIN: Get-PKWindowsEventLog] Get Windows Events matching 'EndTime = 03/22/2022 15:17:48; LogName = Application, Security, System; StartTime = 03/17/2022 15:17:47; Level = 1, 2'
        VERBOSE: [LAPTOP] Searching for event log entries matching filter
        
        MachineName      : LAPTOP.domain.local
        LogName          : System
        TimeCreated      : 2022-03-22 3:08:58 PM
        Id               : 36871
        ProviderName     : Schannel
        Level            : 2
        LevelDisplayName : Error
        Message          : A fatal error occurred while creating a TLS client credential. The internal error state is 10013.

        MachineName      : LAPTOP.domain.local
        LogName          : Application
        TimeCreated      : 2022-03-22 1:12:54 PM
        Id               : 1000
        ProviderName     : Application Error
        Level            : 2
        LevelDisplayName : Error
        Message          : Faulting application name: backgroundTaskHost.exe, version: 10.0.19041.546, time stamp: 0x1d3a15e7
                           Faulting module name: ucrtbase.dll, version: 10.0.19041.789, time stamp: 0x2bd748bf
                           Exception code: 0xc0000409
                           Fault offset: 0x0000000000071208
                           Faulting process id: 0x41fc
                           Faulting application start time: 0x01d83e292d84fde2
                           Faulting application path: C:\WINDOWS\system32\backgroundTaskHost.exe
                           Faulting module path: C:\WINDOWS\System32\ucrtbase.dll
                           Report Id: 256b6816-dde0-41af-8666-e1420af4e233
                           Faulting package full name: Microsoft.SurfaceHub_61.2037.139.0_x64__8wekyb3d8bbwe
                           Faulting package-relative application ID: App

        MachineName      : LAPTOP.domain.local
        LogName          : System
        TimeCreated      : 2022-03-22 10:28:22 AM
        Id               : 12
        ProviderName     : SurfaceTconDriver
        Level            : 2
        LevelDisplayName : Error
        Message          : Surface Tcon Driver TP Write fails, Status = 0xc0000186

        MachineName      : LAPTOP.domain.local
        LogName          : Application
        TimeCreated      : 2022-03-22 10:28:22 AM
        Id               : 65535
        ProviderName     : SurfaceTconHAL
        Level            : 2
        LevelDisplayName : 
        Message          : 

        MachineName      : LAPTOP.domain.local
        LogName          : System
        TimeCreated      : 2022-03-22 10:28:22 AM
        Id               : 13
        ProviderName     : SurfaceTconDriver
        Level            : 2
        LevelDisplayName : Error
        Message          : Surface Tcon Driver TP Read fails, Status = 0xc0000186
        
        <snip>

        VERBOSE: [END: Get-PKWindowsEventLog] Get Windows Events matching 'EndTime = 03/22/2022 15:17:48; LogName = Application, Security, System; StartTime = 03/17/2022 15:17:47; Level = 1, 2'

.EXAMPLE
    PS C:\> Get-PKWindowsEventLog -ComputerName jumpbox.domain.local -Credential (Get-Credential) -StartDate 2021-09-01 -EndDate 2021-09-30  -MaxEvents 5 | Format-Table -AutoSize
    Gets five events from a remote computer with custom start/end dates and the default filters 

        MachineName          LogName     TimeCreated             Id ProviderName      Level LevelDisplayName Message                                                                                      
        -----------          -------     -----------             -- ------------      ----- ---------------- -------                                                                                      
        jumpbox.domain.local Application 2021-09-24 6:08:17 AM 1000 Application Error     2 Error            Faulting application name: mmc.exe, version: 10.0.17763.1697, time stamp: 0x6e7d0aa5...      
        jumpbox.domain.local Application 2021-09-17 7:14:12 PM 1000 Application Error     2 Error            Faulting application name: 14210-CsInstallerService.exe, version: 0.0.0.0, time stamp: 0x6...
        jumpbox.domain.local Application 2021-09-17 2:18:48 AM 1000 Application Error     2 Error            Faulting application name: mmc.exe, version: 10.0.17763.1697, time stamp: 0x6e7d0aa5...      
        jumpbox.domain.local Application 2021-09-16 5:25:56 PM 1000 Application Error     2 Error            Faulting application name: Explorer.EXE, version: 10.0.17763.1911, time stamp: 0xf27ca669... 
        jumpbox.domain.local Application 2021-09-08 3:19:19 PM 1000 Application Error     2 Error            Faulting application name: 14105-CsInstallerService.exe, version: 0.0.0.0, time stamp: 0x6...
 
#>
[CmdletBinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more computer names (default is local computer"
    )]
    [Alias("Identity","Name")]
    [ValidateNotNullOrEmpty()]
    [object[]]$ComputerName,

    [Parameter(
        HelpMessage = "One or more event log names (default is Application, Security, and System; you can select * for all, but this will get extremely chatty!)"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$LogName = @("Application","Security","System"),

    [Parameter(
        HelpMessage = "One or more severity levels for events: Critical, Error, Warning, Information, Warning (default is Critical, Error) "
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Critical","Error","Warning","Information")]
    [string[]]$EventType = @("Critical","Error"),

    [Parameter(
        HelpMessage = "One or more EventIDs (default is all matching events)"
    )]
    [ValidateNotNullOrEmpty()]
    [int]$EventID,

    [Parameter(
        HelpMessage = "Text to search for in events (don't include asterisks; defaults to wildcard search)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Data,

    [Parameter(
        HelpMessage = "Earliest date for filtering matching events (default is past 5 days)"
    )]
    [ValidateNotNullOrEmpty()]
    [datetime]$StartDate = (Get-Date).AddDays(-5),

    [Parameter(
        HelpMessage = "Latest date for filtering matching events (default is now)"
    )]
    [ValidateNotNullOrEmpty()]
    [datetime]$EndDate = (Get-Date),

    [Parameter(
        HelpMessage = "Number of matching events to return (default is all)"
    )]
    [ValidateNotNullOrEmpty()]
    [int]$MaxEvents,

    [Parameter(
        HelpMessage = "Valid credentials on computer (ignored if local; note that authentication failure will still display a generic RPC error)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential]$Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "Connectivity test to run before attempting Get-WinEvent: Ping, RPC test on TCP 135, or None (default is Ping)"
    )]
    [ValidateSet("Ping","RPC","None")]
    [string]$ConnectionTest = "Ping"
)
Begin{
        
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $ScriptName = $MyInvocation.MyCommand.Name
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName 
        # Doing this instead of setting a default in the parameter, because that will always display in the verbose CurrentParams even if we are using pipeline input!
    }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Msg = "Running Get-WinEvent on remote computers requires valid credentials and an enabled inbound firewall rule for 'Remote Event Log Management.'
(Unfortunately, remote authentication failures generate only generic 'RPC connection' failure messages.)
You can enable the firewall rules via Group Policy, or via PowerShell ...it's safest to set this for the Domain profile only.

	PS C:\> Get-NetFirewallRule -Name 'RemoteEventLogSvc-In-TCP' | Set-NetFirewallRule -Enabled True -Profile Domain -Direction Inbound -PassThru`n"
    Write-Verbose $Msg

    # Translate LevelDisplayNames to numbers
    $LookupLevel = @{
        Critical      = 1	 	
        Error 	      = 2
        Warning       = 3
        Informational =	4 	
    }
    $LookupLevel2 = @{
        1 = "Critical"
        2 = "Error"
        3 = "Warning"
        4 = "Informational"
    }

    # Filter hashtable
    $FilterHT = @{}
    $FilterHT = @{
        LogName      = $LogName
        Level        = $LookupLevel[$EventType]
        StartTime    = $StartDate
        EndTime      = $EndDate
    }    
    If ($CurrentParams.EventID) {$FilterHT.Add("ID",$EventID)}
    If ($CurrentParams.Data) {$FilterHT.Add("Data","$Data")}
        
    
    # Splat for Get-WinEvent
    $Param_GWE = @{}
    $Param_GWE = @{
        FilterHashTable = $FilterHT
        Verbose         = $False
        ErrorAction     = "Stop"
        Credential      = $Credential
    }
    If ($CurrentParams.MaxEvents) {$Param_GWE.Add("MaxEvents",$MaxEvents)}

    $HelpTable = "Key name 	Value data type 	Accepts wildcard characters?
LogName 	<String[]> 	Yes
ProviderName 	<String[]> 	Yes
Path 	<String[]> 	No
Keywords 	<Long[]> 	No
ID 	<Int32[]> 	No
Level 	<Int32[]> 	No
StartTime 	<DateTime> 	No
EndTime 	<DateTime> 	No
UserID 	<SID> 	No
Data 	<String[]> 	No
<named-data> 	<String[]> 	No" | ConvertFrom-CSV -Delimiter "`t"
    
    $Select = "MachineName,LogName,TimeCreated,ID,ProviderName,Level,LevelDisplayName,Message" -split(",")

    
    $Activity = "Get Windows Events matching '$(($FilterHT.GetEnumerator() | Foreach-Object {"$($_.Name) = $($_.Value -join(', '))"}) -join("; "))'"
    $Msg = "[BEGIN: $Scriptname] $Activity" 
    Write-Verbose $Msg
}

Process {
    
    $Total = $ComputerName.Count
    $Current = 0        
    Foreach ($Computer in $ComputerName) {
            
        $Current ++
        $Computer = $Computer.Trim()
        $Results = @()
        $ResultMessage = @()

        If ($Computer -match "Localhost|127.0.0.1|^$Env:ComputerName$") {
            [switch]$Continue = $True
        }
        Else {
            Switch ($ConnectionTest) {
                Ping {
                    $Msg = "Pinging computer"
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                    [switch]$Continue = $False
                    If ($Null = Test-Connection $Computer -Quiet -Count 1 -ErrorAction SilentlyContinue) {
                        $Continue = $True
                        $Msg = "Connection succeeded"
                        Write-Verbose "[$Computer] $Msg"
                    }
                    Else {
                        $Msg = "Ping failed"
                        Write-Warning "[$Computer] $Msg"
                        $ResultMessage += $Msg
                        $Output = "" | Select-Object $Select
                        $Output.MachineName = $ComputerName
                        $Output.LogName = $LogName
                        $Output.Message = $ResultMessage
                        $Results += $Output
                    }
                }
                RPC {
                    $Msg = "Testing connection on TCP port 135 for RPC"
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                    [switch]$Continue = $False
                    # Using this method because it's waaaay faster than Test-NetConnection
                    If (($Status = (New-Object System.Net.Sockets.TcpClient).ConnectAsync($Computer,135).Wait(1000))-eq $True) {
                        $Continue = $True
                        $Msg = "Connection succeeded"
                        Write-Verbose "[$Computer] $Msg"
                    }
                    Else {
                        $Msg = "RPC connection test failed or timed out after 1000 milliseconds"
                        Write-Warning "[$Computer] $Msg"
                        $ResultMessage += $Msg
                        Output = "" | Select-Object $Select
                        $Output.MachineName = $ComputerName
                        $Output.LogName = $LogName
                        $Output.Message = $ResultMessage
                        $Results += $Output
                    }
                }
                None {
                    [switch]$Continue = $True
                }
            }
        }
        If ($Continue.IsPresent) {
            
            $Msg = "Searching for event log entries matching filter"
            Write-Verbose "[$Computer] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                        
            Try {
                Get-Winevent -ComputerName $Computer @Param_GWE | Select-Object $Select
            }
            Catch {
                $ResultMessage = $_.Exception.Message
                #$ResultMessage = "Operation failed"
                #If ($ErrorDetails = $_.Exception.Message) {$ResultMessage += " ($ErrorDetails)"}
                $Output = "" | Select-Object $Select
                $Output.MachineName = $ComputerName
                $Output.LogName = $LogName
                $Output.Message = $ResultMessage
                $Results += $Output
            }
        }
        
        Write-Output ($Results | Select-Object $Select)

    } #end foreach 
}
End {

    $Null = Write-Progress -Activity * -Completed
    $Msg = "[END: $Scriptname] $Activity" 
    Write-Verbose $Msg

}
} #end Get-PKWindowsEventLogSettings


$Null = New-Alias Get-PKWinEventLogSettings -Value Get-PKWindowsEventLogSettings -Force -Confirm:$False