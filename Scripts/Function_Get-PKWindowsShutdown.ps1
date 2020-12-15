#Requires -version 3
Function Get-PKWindowsShutdown {
<# 
.SYNOPSIS
    Invokes a scriptblock to query the Windows event log for shutdown/startup events via Get-WinEvent, or last boot time data via Get-WMIObject, interactively or as a PSJob

.DESCRIPTION
    Invokes a scriptblock to query the Windows event log for shutdown/startup events via Get-WinEvent, or last boot time data via Get-WMIObject, interactively or as a PSJob
    Accepts pipeline input
    Tests connectivity to remote computers before invoking scriptblock by default
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsShutdown.ps1
    Created : 2020-07-21
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-07-21 - Created script

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER LastBootTimeOnly
    Display only last boot time data (in computer's local time zone)

.PARAMETER SortOrder
    Output sort order: Ascending or Descending (default is Descending; ignored if LastBootTimeOnly)

.PARAMETER Credential
    Valid credentials on target (default is passthrough)

.PARAMETER Authentication
    WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'Shutdown')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests ignored on local computer)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> Get-PKWindowsShutdown -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                    
        ---              -----                                    
        Verbose          True                                     
        ComputerName     LAPTOP13                          
        LastBootTimeOnly False                                    
        SortOrder        Ascending                                
        Credential       System.Management.Automation.PSCredential
        Authentication   Negotiate                                
        AsJob            False                                    
        JobPrefix        Shutdown                                 
        ConnectionTest   WinRM                                    
        Quiet            False                                    
        ScriptName       Get-PKWindowsShutdown                    
        ScriptVersion    1.0.0                                    
        PipelineInput    False                                    


        [BEGIN: Get-PKWindowsShutdown] Invoke scriptblock to return Windows Event logs events for startup/shutdown/restart

        [LAPTOP13] Invoke command
    
        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : User32
        TimeCreated      : 2020-05-21 11:08:10 AM
        Id               : 1074
        LevelDisplayName : Information
        Message          : The process C:\Windows\System32\RuntimeBroker.exe (LAPTOP13) has initiated the restart of computer LAPTOP13 on behalf of user DOMAIN\jbloggs for the following reason: Other (Unplanned)
                            Reason Code: 0x0
                            Shutdown Type: restart
                            Comment: 

        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : EventLog
        TimeCreated      : 2020-05-21 11:08:52 AM
        Id               : 6006
        LevelDisplayName : Information
        Message          : The Event log service was stopped.

        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : Microsoft-Windows-Kernel-Power
        TimeCreated      : 2020-05-21 11:08:57 AM
        Id               : 109
        LevelDisplayName : Information
        Message          : The kernel power manager has initiated a shutdown transition.
                   
                           Shutdown Reason: Kernel API

        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : Microsoft-Windows-Kernel-General
        TimeCreated      : 2020-05-21 11:09:32 AM
        Id               : 12
        LevelDisplayName : Information
        Message          : The operating system started at system time ‎2020‎-‎05‎-‎21T18:09:31.500000000Z.

        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : Microsoft-Windows-Kernel-Boot
        TimeCreated      : 2020-05-21 11:09:32 AM
        Id               : 20
        LevelDisplayName : Information
        Message          : The last shutdown's success status was true. The last boot's success status was true.

        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : Microsoft-Windows-Kernel-Power
        TimeCreated      : 2020-06-15 8:00:07 AM
        Id               : 41
        LevelDisplayName : Critical
        Message          : The system has rebooted without cleanly shutting down first. This error could be caused if the system stopped responding, crashed, or lost power unexpectedly.

        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : EventLog
        TimeCreated      : 2020-06-15 8:00:22 AM
        Id               : 6005
        LevelDisplayName : Information
        Message          : The Event log service was started.

        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : Microsoft-Windows-UserModePowerService
        TimeCreated      : 2020-06-15 8:00:24 AM
        Id               : 12
        LevelDisplayName : Information
        Message          : Process C:\Windows\System32\Intel\DPTF\esif_uf.exe (process ID:4800) reset policy scheme from {381b4222-f694-41f0-9685-ff5bb260df2e} to {381b4222-f694-41f0-9685-ff5bb260df2e}

        MachineName      : LAPTOP13.domain.local
        LogName          : System
        ProviderName     : Microsoft-Windows-Kernel-General
        TimeCreated      : 2020-06-30 8:05:57 AM
        Id               : 20
        LevelDisplayName : Information
        Message          : The leap second configuration has been updated.
                           Reason: Leap second data initialized from registry during boot
                           Leap seconds enabled: true
                           New leap second count: 0
                           Old leap second count: 0
    
        [END: Get-PKWindowsShutdown] Invoke scriptblock to return Windows Event logs events for startup/shutdown/restart

.EXAMPLE
    PS C:\> $Arr | Get-PKWindowsShutdown -LastBootTimeOnly -Credential $Credential -Quiet 

        
        MachineName  Uptime              LastBootTime          
        -----------  ------              ------------          
        SQLPROD      6.06:04:57.8200231  2020-07-15 11:23:17 AM
        TESTBOX      27.07:57:50.9669382 2020-06-24 9:30:27 AM 

.EXAMPLE
    PS C:\> Get-PKWindowsShutdown -ComputerName "dev-webserver.domain.local","jumpbox.domain.local" -SortOrder Descending -Credential $Credential -AsJob -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                                                           
        ---              -----                                                                           
        ComputerName     {dev-webserver.domain.local, jumpbox.domain.local}
        SortOrder        Descending                                                                      
        Credential       System.Management.Automation.PSCredential                                       
        AsJob            True                                                                            
        Verbose          True                                                                            
        LastBootTimeOnly False                                                                           
        Authentication   Negotiate                                                                       
        JobPrefix        Shutdown                                                                        
        ConnectionTest   WinRM                                                                           
        Quiet            False                                                                           
        ScriptName       Get-PKWindowsShutdown                                                           
        ScriptVersion    1.0.0                                                                           
        PipelineInput    False                                                                           


        [BEGIN: Get-PKWindowsShutdown] Invoke scriptblock to return Windows Event logs events for startup/shutdown/restart (as job)

        [dev-webserver.domain.local] Test WinRM connection
        [dev-webserver.domain.local] Invoke command as PSJob
        [jumpbox.domain.local] Test WinRM connection
        [jumpbox.domain.local] Invoke command as PSJob

        2 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        43     Shutdown_dev... RemoteJob       Completed     True            dev-webserv... ...                      
        45     Shutdown_jum... RemoteJob       Completed     True            jumpbox.gra... ...       

        [END: Get-PKWindowsShutdown] Invoke scriptblock to return Windows Event logs events for startup/shutdown/restart (as job)

        PS C:\> Get-Job 43,45 | Receive-Job | Format-Table -Autosize

        MachineName                             LogName ProviderName                           TimeCreated              Id LevelDisplayName Message                                                                                                          
        -----------                             ------- ------------                           -----------              -- ---------------- -------                                                                                                          
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-07-15 11:24:22 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.19724_none_fa5e641b...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-07-15 11:24:22 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.19724_none_fa5e641b...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-07-15 11:24:09 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.19724_none_fa5e641b...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-07-15 11:24:09 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.19724_none_fa5e641b...
        dev-webserver.domain.local System  EventLog                               2020-07-15 11:23:53 AM 6005 Information      The Event log service was started.                                                                               
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-General       2020-07-15 11:23:17 AM   12 Information      The operating system started at system time ‎2020‎-‎07‎-‎15T18:23:17.487979100Z.                                 
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Boot          2020-07-15 11:23:17 AM   20 Information      The last shutdown's success status was true. The last boot's success status was true.                            
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Power         2020-07-15 11:23:10 AM  109 Information      The kernel power manager has initiated a shutdown transition.                                                    
        dev-webserver.domain.local System  EventLog                               2020-07-15 11:23:05 AM 6006 Information      The Event log service was stopped.                                                                               
        dev-webserver.domain.local System  User32                                 2020-07-15 11:21:39 AM 1074 Information      The process C:\Windows\system32\wbem\wmiprvse.exe (DEV-WEBSERVER) has initiated the restart of computer DEV-...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-06-17 11:28:42 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.19651_none_fa3af193...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-06-17 11:28:42 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.19651_none_fa3af193...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-06-17 11:28:30 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.19651_none_fa3af193...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-06-17 11:28:30 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.19651_none_fa3af193...
        dev-webserver.domain.local System  EventLog                               2020-06-17 11:28:06 AM 6005 Information      The Event log service was started.                                                                               
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-General       2020-06-17 11:26:17 AM   12 Information      The operating system started at system time ‎2020‎-‎06‎-‎17T18:26:17.488052300Z.                                 
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Boot          2020-06-17 11:26:17 AM   20 Information      The last shutdown's success status was true. The last boot's success status was true.                            
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Power         2020-06-17 11:26:04 AM  109 Information      The kernel power manager has initiated a shutdown transition.                                                    
        dev-webserver.domain.local System  EventLog                               2020-06-17 11:25:59 AM 6006 Information      The Event log service was stopped.                                                                               
        dev-webserver.domain.local System  User32                                 2020-06-17 11:23:07 AM 1074 Information      The process C:\Windows\system32\wbem\wmiprvse.exe (DEV-WEBSERVER) has initiated the restart of computer DEV-...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-05-20 11:49:51 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.18384_none_fa1d93c3...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-05-20 11:49:51 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.18384_none_fa1d93c3...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-05-20 11:49:37 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.18384_none_fa1d93c3...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-05-20 11:49:37 AM   12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.3.9600.18384_none_fa1d93c3...
        dev-webserver.domain.local System  EventLog                               2020-05-20 11:49:14 AM 6005 Information      The Event log service was started.                                                                               
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-General       2020-05-20 11:45:42 AM   12 Information      The operating system started at system time ‎2020‎-‎05‎-‎20T18:45:42.486764000Z.                                 
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Boot          2020-05-20 11:45:42 AM   20 Information      The last shutdown's success status was true. The last boot's success status was true.                            
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Power         2020-05-20 11:45:34 AM  109 Information      The kernel power manager has initiated a shutdown transition.                                                    
        dev-webserver.domain.local System  EventLog                               2020-05-20 11:45:27 AM 6006 Information      The Event log service was stopped.                                                                               
        dev-webserver.domain.local System  User32                                 2020-05-20 11:40:06 AM 1074 Information      The process C:\Windows\system32\wbem\wmiprvse.exe (DEV-WEBSERVER) has initiated the restart of computer DEV-...
        dev-webserver.domain.local System  EventLog                               2020-04-14 9:42:19 AM  6005 Information      The Event log service was started.                                                                               
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-General       2020-04-14 9:38:08 AM    12 Information      The operating system started at system time ‎2020‎-‎04‎-‎14T16:38:08.487306000Z.                                 
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Boot          2020-04-14 9:38:08 AM    20 Information      The last shutdown's success status was true. The last boot's success status was true.                            
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Power         2020-04-14 9:37:55 AM   109 Information      The kernel power manager has initiated a shutdown transition.                                                    
        dev-webserver.domain.local System  EventLog                               2020-04-14 9:37:50 AM  6006 Information      The Event log service was stopped.                                                                               
        dev-webserver.domain.local System  User32                                 2020-04-14 9:37:44 AM  1074 Information      The process msiexec.exe has initiated the restart of computer DEV-WEBSERVER on behalf of user NT AUTHORITY\S...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-04-14 9:37:04 AM    12 Information      Process C:\Windows\System32\msiexec.exe (process ID:3120) reset policy scheme from {A1841308-3541-4FAB-BC81-F7...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-04-14 9:37:04 AM    12 Information      Process C:\Windows\System32\msiexec.exe (process ID:3120) reset policy scheme from {8C5E7FDA-E8BF-4A96-9A85-A6...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-04-14 9:37:04 AM    12 Information      Process C:\Windows\System32\msiexec.exe (process ID:3120) reset policy scheme from {381B4222-F694-41F0-9685-FF...
        dev-webserver.domain.local System  Microsoft-Windows-UserModePowerService 2020-04-14 9:37:04 AM    12 Information      Process C:\Windows\System32\msiexec.exe (process ID:3120) reset policy scheme from {8C5E7FDA-E8BF-4A96-9A85-A6...
        dev-webserver.domain.local System  EventLog                               2020-03-18 3:01:16 AM  6005 Information      The Event log service was started.                                                                               
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Boot          2020-03-18 2:58:13 AM    20 Information      The last shutdown's success status was true. The last boot's success status was true.                            
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-General       2020-03-18 2:58:13 AM    12 Information      The operating system started at system time ‎2020‎-‎03‎-‎18T09:58:13.488309900Z.                                 
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-General       2020-01-31 11:22:53 AM   12 Information      The operating system started at system time ‎2020‎-‎01‎-‎31T19:22:53.488005500Z.                                 
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Boot          2020-01-31 11:22:53 AM   20 Information      The last shutdown's success status was true. The last boot's success status was true.                            
        dev-webserver.domain.local System  Microsoft-Windows-Kernel-Power         2020-01-31 11:22:45 AM  109 Information      The kernel power manager has initiated a shutdown transition.                                                    
        dev-webserver.domain.local System  EventLog                               2020-01-31 11:22:40 AM 6006 Information      The Event log service was stopped.                                                                               
        dev-webserver.domain.local System  User32                                 2020-01-31 11:22:28 AM 1074 Information      The process C:\Windows\explorer.exe (DEV-WEBSERVER) has initiated the restart of computer DEV-WEBSERVER on    ...
        dev-webserver.domain.local System  User32                                 2020-01-31 11:22:28 AM 1074 Information      The process explorer.exe has initiated the restart of computer DEV-WEBSERVER on behalf of user DOMAIN\jbloggs ...
        jumpbox.domain.local       System  Microsoft-Windows-UserModePowerService 2020-06-24 9:34:53 AM    12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.2.9200.23060_none_62c3f838...
        jumpbox.domain.local       System  Microsoft-Windows-UserModePowerService 2020-06-24 9:34:53 AM    12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.2.9200.23060_none_62c3f838...
        jumpbox.domain.local       System  Microsoft-Windows-UserModePowerService 2020-06-24 9:34:27 AM    12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.2.9200.23060_none_62c3f838...
        jumpbox.domain.local       System  Microsoft-Windows-UserModePowerService 2020-06-24 9:34:27 AM    12 Information      Process C:\Windows\WinSxS\amd64_microsoft-windows-servicingstack_31bf3856ad364e35_6.2.9200.23060_none_62c3f838...
        jumpbox.domain.local       System  EventLog                               2020-06-24 9:33:18 AM  6005 Information      The Event log service was started.                                                                               
        jumpbox.domain.local       System  Microsoft-Windows-Kernel-Boot          2020-06-24 9:30:28 AM    20 Information      The last shutdown's success status was true. The last boot's success status was true.                            
        jumpbox.domain.local       System  Microsoft-Windows-Kernel-General       2020-06-24 9:30:28 AM    12 Information      The operating system started at system time ‎2020‎-‎06‎-‎24T16:30:27.492122100Z.                                 
        jumpbox.domain.local       System  Microsoft-Windows-Kernel-Power         2020-06-24 9:30:08 AM   109 Information      The kernel power manager has initiated a shutdown transition.                                                    
        jumpbox.domain.local       System  EventLog                               2020-06-24 9:30:03 AM  6006 Information      The Event log service was stopped.                                                                               
        jumpbox.domain.local       System  User32                                 2020-06-24 9:25:24 AM  1074 Information      The process C:\Windows\system32\wbem\wmiprvse.exe (jumpbox) has initiated the restart of computer JUMPB...
        jumpbox.domain.local       System  EventLog                               2020-05-27 9:13:13 AM  6005 Information      The Event log service was started.                                                                               
        jumpbox.domain.local       System  Microsoft-Windows-Kernel-General       2020-05-27 9:12:20 AM    12 Information      The operating system started at system time ‎2020‎-‎05‎-‎27T16:12:19.493912600Z.                                 
        jumpbox.domain.local       System  Microsoft-Windows-Kernel-Boot          2020-05-27 9:12:20 AM    20 Information      The last shutdown's success status was true. The last boot's success status was true.                            
        jumpbox.domain.local       System  Microsoft-Windows-Kernel-Power         2020-05-27 9:11:40 AM   109 Information      The kernel power manager has initiated a shutdown transition.                                                    
        jumpbox.domain.local       System  EventLog                               2020-05-27 9:11:35 AM  6006 Information      The Event log service was stopped.                                                                               
        jumpbox.domain.local       System  User32                                 2020-05-27 9:11:26 AM  1074 Information      The process C:\Windows\system32\wbem\wmiprvse.exe (jumpbox) has initiated the restart of computer JUMPB...
#> 

[CmdletBinding(
    DefaultParameterSetName = "Default",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more computer names (default is local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [object[]]$ComputerName,

    [Parameter(
        HelpMessage = "Display only last boot time data (in computer's local time zone)"
    )]
    [Switch]$LastBootTimeOnly,

    [Parameter(
        HelpMessage = "Output sort order: Ascending or Descending (default is Descending; ignored if LastBootTimeOnly)"
    )]
    [ValidateSet('Ascending','Descending')]
    [string]$SortOrder = "Descending",

    [Parameter(
        HelpMessage = "Valid credentials on target (default is passthrough)"
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
        HelpMessage = "Prefix for job name (default is 'Shutdown')"
    )]
    [String] $JobPrefix = "Shutdown",

    [Parameter(
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests are ignored on local computer)"
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
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName
    }
    $CurrentParams.Add("ScriptName",$Scriptname)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | Out-String )"

    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        Param($SortOrder,$LastBootTimeOnly)
        
        If ($LastBootTimeOnly.IsPresent) {
            
            Try {
                $Output = Get-WmiObject win32_operatingsystem -ErrorAction Stop | 
                    Select-Object @{N='ComputerName';E={$Env:ComputerName}},@{N=’Uptime’;E={([DateTime]::Now – $_.ConvertToDateTime($_.LastBootUpTime))}},@{N=’LastBootTime’;E={$_.ConverttoDateTime($_.lastbootuptime)}}
            }
            Catch {
                $Output = New-Object PSObject -Property @{
                    MachineName  = $Env:ComputerName
                    Uptime       = "Error"
                    LastBootTime = $_.Exception.Message
                } | Select-Object MachineName,Uptime,LastBootTime
            }
        }
        Else {
            $Select = "MachineName,LogName,ProviderName,TimeCreated,ID,LevelDisplayName,Message" -Split(",")
            $IDs = "1074,1076,6006,109,20,12,6005,41"  -Split(",")
        
            $SortParam = @{
                Property = "TimeCreated"
            }
            If ($SortOrder -eq "Descending") {
                $SortParam.Add("Descending",$True)
            }
        
            Try {
                $Output = Get-Winevent -FilterHashtable @{logname = 'System'; id = $IDs} -ErrorAction Stop | 
                    Select-Object @{N='ComputerName';E={$Env:ComputerName}},LogName,ProviderName,TimeCreated,ID,LevelDisplayName,Message | Sort-Object @SortParam 
            }
            Catch {
                $Output = New-Object PSObject -Property @{
                    ComputerName     = $Env:ComputerName
                    LogName          = "System"
                    ProviderName     = "Error"
                    TimeCreated      = "Error"
                    ID               = "Error"
                    LevelDisplayName = "Error"
                    Message          = $_.Exception.Message
                } | Select-Object ComputerName,LogName,ProviderName,TimeCreated,ID,LevelDisplayName,Message
            }
        }

        Write-Output $Output

    } #end scriptblock

    #endregion Scriptblock for Invoke-Command

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

    # Function to write an error as a string (no stacktrace), or an error, and options for prefix to string
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        If (-not $Quiet.IsPresent) {
            $Host.UI.WriteErrorLine("$Message")
        }
        Else {Write-Error "$Message"}
    }
    # Function to write a warning, with any error data, and options for prefix to string
    Function Write-MessageWarning {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Warning $Msg
    }

    # Function to write a verbose message, collecting error data
    Function Write-MessageVerbose {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Verbose $Msg
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
        }
        Try {
            If (Test-WSMan @Param_WSMAN) {$True}
            Else {$False}
        }
        Catch {$False}
    }

    # Function to test ping connectivity
    Function Test-Ping{
        Param($Computer)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
    }

    #endregion Functions

    #region Splats

    # Splat for Write-Progress
    $Activity = "Invoke scriptblock"
    Switch ($LastBootTimeOnly) {
        $True {$Activity += " to return last boot time data"}
        $False {$Activity += " to return Windows Event logs events for startup/shutdown/restart"}
    }
    If ($AsJob.IsPresent) {
        $Activity = "$Activity (as job)"
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
        ArgumentList   = $SortOrder,$LastBootTimeOnly
        ScriptBlock    = $ScriptBlock
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
    $Param_IC_Local = @{}
    $Param_IC_Local = @{
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $SortOrder,$LastBootTimeOnly
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
    "[BEGIN: $ScriptName] $Activity" | Write-MessageInfo -FGColor Yellow -Title


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
        
        $Current ++ 
        $Param_WP.PercentComplete = ($Current/$Total* 100)
        $Param_WP.Status = $Computer
        
        [switch]$Continue = $False
        [bool]$IsLocal = $True

        If ($Computer -in ($Env:ComputerName,"localhost","127.0.0.1")) {
            $IsLocal = $True
            $Continue = $True
        }
        Else {
            Switch ($ConnectionTest) {
                Default {$Continue = $True}
                Ping {
                    $Msg = "Ping computer"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                        If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "Ping failure"
                            "[$Computer] $Msg" | Write-MessageError
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                    }
                }
                WinRM {
                    $Msg = "Test WinRM connection"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                        If ($Null = Test-WinRM -Computer $Computer) {
                            $Continue = $True
                        }
                        Else {
                            $Msg = "WinRM failure"
                            "[$Computer] $Msg" | Write-MessageError
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                    }
                }        
            }
        } #end if not local

        If ($Continue.IsPresent) {
            
            $ConfirmMsg = "`n`n`t$Activity`n`n"
            If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
                Try {
                    $Msg = "Invoke command"
                    If ($AsJob.IsPresent) {$Msg += " as PSJob"}
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

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
                "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

    If ($AsJob.IsPresent -and ($Jobs.Count -gt 0)) {

        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output`n"
            "$Msg" | Write-MessageInfo -FGColor White -Title
            $Jobs | Get-Job
            
        }
        Else {
            $Msg = "No jobs created"
            $Msg | Write-MessageError
        }
    } #end if AsJob

    "[END: $ScriptName] $Activity" | Write-MessageInfo -FGColor Yellow -Title


}

} # end Get-PKWindowsShutdown

$Null = New-Alias Get-PKWindowsBootInfo -Value Get-PKWindowsShutdown -Description "for guessability" -Force -Confirm:$False
$Null = New-Alias Get-PKWindowsRestart -Value Get-PKWindowsShutdown -Description "for guessability" -Force -Confirm:$False