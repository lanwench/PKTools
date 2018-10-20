#requires -Version 3
Function Get-PKWindowsReport { 
<# 
.SYNOPSIS 
        This function will collect various data elements from a local or remote computer. 
.DESCRIPTION 
        This function was inspired by Get-ServerInfo a custom function written by Jason Walker  
        and the PSInfo Sysinternals Tool written by Mark Russinovich.  It will collect a plethora 
        of data elements that are important to a Microsoft Windows System Administrator.  The  
        function will run locally, run without the -ComputerName Parameter, or connect remotely  
        via the -ComputerName Parameter.  This function will return objects that you can interact  
        with, however, due to the fact that multiple custom objects are returned, when piping the  
        function to Get-Member, it will only display the first object, unless you run the following;  
        "Get-ComputerInfo | Foreach-Object {$_ | Get-Member}".  This function is currently in beta.   
        Also remember that you have to dot source the ".ps1" file in order to load it into your  
        current PowerShell console: ". .\Get-ComputerInfo.ps1"  Then it can be run as a "cmdlet"  
        aka "function".  Reminder: In it's current state, this function's output is intended for  
        the console, in other words the data does not export very well, unless the Foreach-Object  
        technique is used above.  This is something that may come in a future release or a simplied  
        version. 
.PARAMETER ComputerName 
        A single Computer or an array of computer names.  The default is localhost ($env:COMPUTERNAME). 
.EXAMPLE 
        PS D:\> Get-ComputerInfo -ComputerName LAP01 
 
        Computer            : LAP01 
        Domain              : WILHITE.DS 
        OperatingSystem     : Microsoft Windows 8 Enterprise 
        OSArchitecture      : 64-bit 
        BuildNumber         : 9200 
        ServicePack         : 0 
        Manufacturer        : Microsoft Corporation 
        Model               : Virtual Machine 
        SerialNumber        : 123456789 
        Processor           : Intel(R) Xeon(R) CPU E5-2670 0 @ 2.60GHz 
        LogicalProcessors   : 2 
        PhysicalMemory      : 8192 
        OSReportedMemory    : 8187 
        PAEEnabled          : False 
        InstallDate         : 8/16/2012 7:44:31 PM 
        LastBootUpTime      : 3/6/2013 8:44:57 AM 
        UpTime              : 04:32:17.1965808 
        RebootPending       : False 
        RebootPendingKey    : False 
        CBSRebootPending    : False 
        WinUpdRebootPending : False 
 
        NetAdapterName  : Microsoft Virtual Machine Bus Network Adapter 
        NICManufacturer : Microsoft 
        DHCPEnabled     : False 
        MACAddress      : 00:15:5D:01:05:5D 
        IPAddress       : {10.1.1.22} 
        IPSubnetMask    : {255.255.255.0} 
        DefaultGateway  : {10.1.1.1} 
        DNSServerOrder  : {10.1.1.10, 10.1.1.11} 
        DNSSuffixSearch : {wilhite.ds} 
        PhysicalAdapter : True 
        Speed           : 10000 Mbit 
 
        DeviceID    : C: 
        VolumeName  : OS 
        VolumeDirty : False 
        Size        : 24.90 GB 
        FreeSpace   : 14.21 GB 
        PercentFree : 57.05 % 
 
        DeviceID    : D: 
        VolumeName  : Data 
        VolumeDirty : False 
        Size        : 100.00 GB 
        FreeSpace   : 61.66 GB 
        PercentFree : 61.66 % 
.LINK 
        Registry Class 
        http://msdn.microsoft.com/en-us/library/microsoft.win32.registry.aspx 
 
        Win32_BIOS 
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa394077(v=vs.85).aspx 
 
        Win32_ComputerSystem 
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa394102(v=vs.85).aspx 
 
        Win32_OperatingSystem 
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa394239(v=vs.85).aspx 
 
        Win32_NetworkAdapter 
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa394216(v=vs.85).aspx 
 
        Win32_NetworkAdapterConfiguration 
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa394217(v=vs.85).aspx 
 
        Win32_Processor 
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa394373(v=vs.85).aspx 
 
        Win32_PhysicalMemory 
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa394347(v=vs.85).aspx 
 
        Win32_LogicalDisk 
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa394173(v=vs.85).aspx 
 
        Component-Based Servicing 
        http://technet.microsoft.com/en-us/library/cc756291(v=WS.10).aspx 
 
        PendingFileRename/Auto Update: 
        http://support.microsoft.com/kb/2723674 
        http://technet.microsoft.com/en-us/library/cc960241.aspx 
        http://blogs.msdn.com/b/hansr/archive/2006/02/17/patchreboot.aspx 
 
        SCCM 2012/CCM_ClientSDK: 
        http://msdn.microsoft.com/en-us/library/jj902723.aspx 
.NOTES 
        Author:    Brian C. Wilhite 
        Email:     bwilhite1@carolina.rr.com 
        Date:      03/31/2012 
        RevDate:   08/29/2012 
        PoShVer:   2.0/3.0 
        ScriptVer: 0.86 (Beta) 
        0.86 - Code clean-up, now a bit easier to read 
                Added several PendingReboot properites 
                RebootPendingKey - Shows contents of files pending rename 
                CBSRebootPending - Component-Based Servicing, see link above 
                WinUpdRebootPending - Pending Reboot due to Windows Update 
                Added PAEEnabled Property 
        0.85 - Now reports LogicalProcessors & Domain (2K3/2K8) 
                Better PendingReboot support for Windows 2008+ 
                Minor Write-Progress Changes 
                    
#> 
 
[CmdletBinding()] 
param( 
    [Parameter(
        Position=0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )] 
    [Alias("CN","Computer")] 
    [String[]]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(
        HelpMessage = "Valid credential on target Windows computer (for WMI query)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty ,

    [Parameter(
        HelpMessage = "Return network and disk data as collections or as a summary"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Collections","Summary")]
    [string]$OutputType = "Collections"
) 
 
Begin  { 

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # How did we get here
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("VM")) -and (-not $ComputerName)
    $Source = $PSCmdlet.ParameterSetName

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # For progress
    $i=0 

    #region Selection arrays for display order
     
    # $CompInfo object
    $Select_CompInfo = @( 
        "ComputerName"
        "Domain"
        "OperatingSystem"
        "OSArchitecture"
        "BuildNumber"
        "ServicePack"
        "IsVirtual"
        "Manufacturer"
        "Model"
        "SerialNumber"
        "Processor"
        "LogicalProcessor"
        "PhysicalMemoryGB"
        "OSReportedMemoryGB"
        "HardDriveCount"
        "HardDriveInfo"
        "TotalGB"
        "TotalFree"
        "NetAdapterCount"
        "NetAdapterInfo"
        "MACAddress"
        "IPAddress"
        "PAEEnabled"
        "ForwardDNSLookup"
        "ReverseDNSLookup"
        "InstallDate"
        "LastBootUpTime"
        "UpTime"
        "RebootPending"  
    ) 
     
    # $NetInfo object
    $Select_NetInfo = @( 
        "NetConnectionID"
        "NICName" 
        "NICManufacturer" 
        "DHCPEnabled" 
        "MACAddress" 
        "IPAddress" 
        "IPSubnetMask" 
        "IPv4SubnetPrefix"
        "DefaultGateway" 
        "DNSServerOrder" 
        "DNSSuffixSearch" 
        "PhysicalAdapter" 
        "Speed" 
    ) 
     
    # $VolInfo object
    $Select_VolInfo = @( 
        "DeviceID" 
        "VolumeName" 
        "VolumeDirty" 
        "SizeGB" 
        "FreeGB" 
        "PercentFree" 
    )
    
    #endregion Selection arrays for display order

    # Scriptblock for Invoke-Command
    $Scriptblock = {
    
        $PendingOps = 0
        Try {
            If (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA SilentlyContinue) {$PendingOps ++}
            If (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA SilentlyContinue) {$PendingOps ++}
            If (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA SilentlyContinue) {$PendingOps ++}
            If ($PendingOps.Count -gt 0) {$True} 
            Else {$False}
        }
        Catch {
            $_.Exception.Message
        }
    }

    #endregion properties
    
    #region Inner functions

    # Test forward/reverse DNS lookups
    Function Test-VM ($BIOS,$CompSystem) {
        $DCVM = "Error"
        if($BIOS.Version -match "VRTUAL") {$DCVM = "Virtual - Hyper-V"}
        elseif($BIOS.Version -match "A M I") {$DCVM = "Virtual - Virtual PC"}
        elseif($BIOS.Version -like "*Xen*") {$DCVM = "Virtual - Xen"}
        elseif($BIOS.SerialNumber -like "*VMware*") {$DCVM = "Virtual - VMWare"}
        elseif($CompSystem.manufacturer -like "*Microsoft*" -and ($CompSystem.Model -notmatch "surface")) {$DCVM = "Virtual - Hyper-V"}
        elseif($CompSystem.manufacturer -like "*VMWare*") {$DCVM = "Virtual - VMWare"}
        elseif($CompSystem.model -like "*Virtual*") {$DCVM = "Virtual"}
        else {$DCVM = "Physical"}
        #Write-Output $DCVM
        If ($DCVM -match "Physical") {$False}
        ElseIf ($DCVM -match "Virtual")  {$True}
        Else {$DVCM}
    } #end Test-VM

    # Test forward/reverse DNS lookups, trapping/hiding errors
    Function Test-DNS{
        [Cmdletbinding()]
        Param($Name,$IP)
        If ($Name) {
            Try {
                [Net.Dns]::GetHostEntry($Name)
            }
            Catch {}
        }
        Elseif ($IP) {
            Try {
                [Net.Dns]::GetHostByAddress($IP)
            }
            Catch {}
        }
        
    } #end Test-DNS

    # Convert subnet mask to CIDR notation
    function Convert-Mask([string] $Mask){
        # https://d-fens.ch/2013/11/01/nobrainer-using-powershell-to-convert-an-ipv4-subnet-mask-length-into-a-subnet-mask-address/
        $result = 0 
        # ensure we have a valid IP address
        [IPAddress] $ip = $Mask
        $octets = $ip.IPAddressToString.Split('.')
        foreach($octet in $octets){
            while(0 -ne $octet) {
                $octet = ($octet -shl 1) -band [byte]::MaxValue
                $result++
            }
        }
        return $result
    } #end Convert-Mask

    #endregion Inner functions
        
    #region Splats
    
    # Write-Progress
    $Activity = "Getting Windows report data"
    $Param_WP1 = @{}
    $Param_WP1 = @{ 
        Id               = 1 
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"  
        PercentComplete  = $Null
    }     
    
    # Write-Progress        
    $Param_WP2 = @{}
    $Param_WP2 = @{ 
        ParentId        = 1  
        Activity        = $Activity 
        Status          = $Null 
        PercentComplete = $Null
    }

    # Get-WMIObject
    $Param_GWMI = @{}
    $Param_GWMI = @{
        ComputerName = $Null
        Class        = $Null
        Credential   = $Credential
        ErrorAction  = "Stop"
        Verbose      = $False
    } 
    
    # Invoke-Command
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName = $Null
        Scriptblock  = $Scriptblock
        Credential   = $Credential
        ErrorAction  = "Stop"
        Verbose      = $False
    }   

    #endregion Splats

}
 
Process  { 

    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        Try  { 
            
            $Current ++

            # Splats & progress bar
            $Param_WP1.CurrentOperation =$Param_GWMI.ComputerName = $Param_IC.Computername = $Param_WP1.CurrentOperation = $Computer
            $Param_WP1.PercentComplete = $Current/$Total * 100
            Write-Progress @Param_WP1 
            
            # Progress bar counters 
            $n,$d = 0,11 

            #region WMI data

            # 1: Processor
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_Processor"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2
            $WMI_PROC = Get-WmiObject @Param_GWMI
            $n++ 
                         
            # 2: BIOS
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_BIOS"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $WMI_BIOS = Get-WmiObject @Param_GWMI
            $n++

            # 3: ComputerSystem
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_ComputerSystem"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $WMI_CS = Get-WmiObject @Param_GWMI
            $n++

            # 4: OperatingSystem
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_OperatingSystem"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $WMI_OS = Get-WmiObject @Param_GWMI
            $n++

            # 5: RAM
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_PhysicalMemory"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $WMI_PM = Get-WmiObject @Param_GWMI
            $n++
 
            # 6: LogicalDisk
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_LogicalDisk"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $WMI_LD = Get-WmiObject @Param_GWMI -Filter "DriveType = '3'"
            $n++

            # 7: PhysicalDisk
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_DiskDrive"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $WMI_PD = Get-WmiObject @Param_GWMI 
            $n++

            # 8: Network adapter
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_NetworkAdapter"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $WMI_NA = Get-WmiObject @Param_GWMI
            $n++
            
            # 9: Network adapter config
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_GWMI.Class = "Win32_NetworkAdapterConfiguration"
            $Param_WP2.Activity = "Collecting data: $($Param_GWMI.Class)"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $WMI_NAC = Get-WmiObject @Param_GWMI -Filter "IPEnabled=$true"
            $n++

            # OS build
            [int]$WinBuild = $WMI_OS.BuildNumber 
            
            # Logical processors
            $OSArchitecture = "(not available)"
            If ($WinBuild -ge 6001)  { 
                $OSArchitecture = $WMI_OS.OSArchitecture 
                $LogicalProcs   = $WMI_CS.NumberOfLogicalProcessors 
            }
            Else  {
                $LogicalProcs = 1
                If ($WMI_PROC.Count -gt 1) { 
                    $LogicalProcs = $WMI_PROC.Count 
                } 
            }

            #endregion WMI data

            #region Misc

            # 10: Registry (pending reboot)
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_WP2.Activity = "Collecting data: Registry query"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $RegCon = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]"LocalMachine",$Computer) 
            $n++
            
            $RebootPending = "(not available)" 
            If ($WinBuild-ge 6001) {
                $Param_IC.ComputerName = $Computer
                $RebootPending = Invoke-Command @Param_IC
            }

            # 11: Architecture (p or v)
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_WP2.Activity = "Collecting data: Computer architecture"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $IsVirtual = Test-VM -BIOS $WMI_BIOS -CompSystem $WMI_CS
            $n++
            
            # 12: DNS
            $PercentComplete = [math]::Round(($n/$d * 100))
            $Param_WP2.Activity = "Collecting data: DNS records"
            $Param_WP2.Status = "Percent complete: $PercentComplete %"
            $Param_WP2.PercentComplete = $PercentComplete
            Write-Progress @Param_WP2 
            $Forward = Test-DNS -Name $Computer
            $Reverse = Test-DNS -IP ($WMI_NAC.IPAddress | Where-Object {$_ -notmatch ":"})
            $StrForward = "$($Forward.HostName) -> $($Forward.AddressList.IPAddressToString)"
            $StrReverse = "$($Reverse.AddressList.IPAddressToString) -> $($Reverse.HostName)"
            $n++

            #endregion Misc

            #region Calculated properties

            #Calculating Memory, Converting InstallDate, LastBootTime, Uptime. 
            [int]$Memory  = ($WMI_PM | Measure-Object -Property Capacity -Sum).Sum / 1GB 
            $InstallDate  = ([WMI]'').ConvertToDateTime($WMI_OS.InstallDate) 
            $LastBootTime = ([WMI]'').ConvertToDateTime($WMI_OS.LastBootUpTime) 
            $UpTime       = New-TimeSpan -Start $LastBootTime -End (Get-Date) 
            $StrUptime = "$($Uptime.Days)d $($Uptime.Hours)h $($Uptime.Minutes)m $($Uptime.Seconds)s"
             
            #PAEEnabled is only valid on x86 systems, setting value to false first. 
            $PAEEnabled = $false 
            If ($WMI_OS.PAEEnabled)  { 
                $PAEEnabled = $true 
            } 

            #endregion Calculated properties

            #region Output objects

            #region Network object
             
            #There may be multiple NICs that have IPAddresses, hence the Foreach loop. 
            [array]$ObjNetwork = Foreach ($NAC in $WMI_NAC)  { 
                #Getting properties from $WMI_NA that correlate to the matched Index, this is faster than using $WMI_NAC.GetRelated('Win32_NetworkAdapter').  
                $NetAdap = $WMI_NA | Where-Object {$NAC.Index -eq $_.Index} 
                 
                #Since there are properties that are exclusive to 2K8+ marking "Unaval" for computers below 2K8. 
                $PhysAdap = $Speed = "**Unavailable**"

                If ($WinBuild -ge 6001) { 
                    $PhysAdap = $NetAdap.PhysicalAdapter 
                    $Speed    = "{0:0} Mbit" -f $($NetAdap.Speed / 1000000) 
                }
                $Prefix = Convert-Mask -Mask $($NAC.IPSubnet | Where-Object {$_ -match "\d\."})
                
                #Creating the $NetInfo Object 
                New-Object PSObject -Property @{ 
                    NetConnectionID  = $NetAdap.NetConnectionID
                    NICName          = $NetAdap.Name 
                    NICManufacturer  = $NetAdap.Manufacturer 
                    DHCPEnabled      = $NAC.DHCPEnabled 
                    MACAddress       = $NAC.MACAddress 
                    IPAddress        = $NAC.IPAddress 
                    IPSubnetMask     = $NAC.IPSubnet 
                    IPv4SubnetPrefix = $Prefix
                    DefaultGateway   = $NAC.DefaultIPGateway 
                    DNSServerOrder   = $NAC.DNSServerSearchOrder 
                    DNSSuffixSearch  = $NAC.DNSDomainSuffixSearchOrder 
                    PhysicalAdapter  = $PhysAdap 
                    Speed            = $Speed 
                } | Select-Object $Select_NetInfo 
 
            }#End Foreach ($NAC in $WMI_NAC) 
            
            #endregion Network object
            
            #region Disk object
               
            # Creating the $VolInfo Object  - there may be multiple Volumes, hence the Foreach loop.   
            [array]$ObjDisk = Foreach ($Volume in $WMI_LD)  { 
                
                New-Object PSObject -Property @{ 
                    DeviceID    = $Volume.DeviceID 
                    VolumeName  = $Volume.VolumeName 
                    VolumeDirty = $Volume.VolumeDirty 
                    SizeGB      = [math]::Round($Volume.Size /1GB)
                    FreeGB      = [math]::round($Volume.Freespace/1GB)
                    PercentFree = [math]::Round(($Volume.FreeSpace/1GB) / ($Volume.Size/1GB) * 100)
                    #Size        = $("{0:F} GB" -f $($Volume.Size / 1GB)) 
                    #FreeSpace   = $("{0:F} GB" -f $($Volume.FreeSpace / 1GB)) 
                    #PercentFree = $("{0:P}" -f $($Volume.FreeSpace / $Volume.Size)) 
                    } | Select-Object $Select_VolInfo
 
            }#End Foreach ($Volume in $WMI_LD) 

            #endregion Disk object

            #region Compinfo object 
            $ObjComputer = New-Object PSObject -Property @{ 
                ComputerName        = $WMI_CS.Name 
                Domain              = $WMI_CS.Domain.ToUpper() 
                OperatingSystem     = $WMI_OS.Caption 
                OSArchitecture      = $OSArchitecture 
                BuildNumber         = $WinBuild 
                ServicePack         = $WMI_OS.ServicePackMajorVersion 
                IsVirtual           = $IsVirtual
                Manufacturer        = $WMI_CS.Manufacturer 
                Model               = $WMI_CS.Model 
                SerialNumber        = $WMI_BIOS.SerialNumber 
                Processor           = ($WMI_PROC | Select-Object -ExpandProperty Name -First 1) 
                LogicalProcessor    = $LogicalProcs 
                PhysicalMemoryGB    = $Memory
                OSReportedMemoryGB  = [int]$($WMI_CS.TotalPhysicalMemory / 1GB) 
                HardDriveCount      = $ObjDisk.Count
                HardDriveInfo       = $Null
                TotalGB             = [math]::round(($ObjDisk.SizeGB | Measure-Object -Sum).Sum)
                TotalFree           = [math]::round(($ObjDisk.FreeGB | Measure-Object -Sum).Sum) 
                NetAdapterCount     = $ObjNetwork.Count
                NetAdapterInfo      = $Null
                MACAddress          = $ObjNetwork.MACAddress
                IPAddress           = $ObjNetwork.IPAddress
                PAEEnabled          = $PAEEnabled 
                ForwardDNSLookup    = $StrForward
                ReverseDNSLookup    = $StrReverse
                InstallDate         = $InstallDate 
                LastBootUpTime      = $LastBootTime 
                UpTime              = $StrUpTime 
                RebootPending       = $RebootPending  
            }

            $TotalGB = ($ObjDisk.SizeGB | Measure-Object -Sum).Sum
            $TotalFree = ($ObjDisk.FreeGB | Measure-Object -Sum).Sum

            Switch ($OutputType) {
                Collections {
                    $ObjComputer.HardDrive = $ObjDisk
                    $ObjComputer.NetworkAdapter = $ObjNetwork
                }               
                Summary {
                    $DiskSummary = ($ObjDisk | ForEach-Object {"$($_.DeviceID) [$($_.VolumeName)] $($_.SizeGB)GB ($($_.PercentFree)% free)"}) -join(", ")
                    $ObjComputer.HardDrive = $DiskSummary
                    
                    $NetworkSummary = ($ObjNetwork | 
                        ForEach-Object {
                            "[$($_.NetConnectionID)] MACAddress: $($_.MACAddress); IPv4address: $(($_.IPAddress | Where-Object {$_ -notmatch ":"}) -join(", "))/$Prefix, Gateway: $($_.DefaultGateway); DNS: $($_.DNSServerOrder -join(", ")); DHCPEnabled: $($_.DHCPEnabled) "}
                    ) -join(", ")
                    $ObjComputer.NetworkAdapter = $NetworkSummary
                }
            }  
            
            #endregion Compinfo object
            
            Write-Output ($ObjComputer  | Select-Object $Select_CompInfo )  
        }
        Catch { 
            Throw $_.Exception.Message
        }

    } # end for each computer

}#End Process 
   
End { 
    Write-Progress -Activity $Activity -Completed
    Write-Progress -Activity $Param_WP2.Activity -Completed
    $ErrorActionPreference = $TempErrAct 
} 
 
}#End Function Get-ComputerInfo
