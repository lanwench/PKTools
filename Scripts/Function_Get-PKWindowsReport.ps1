#Requires -version 3
Function Get-PKWindowsReport {
<# 
.SYNOPSIS
    Returns Windows computer report data, interactively or as a PSJob

.DESCRIPTION
    Returns Windows computer report data, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsReport.ps1
    Created : 2018-09-11
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-03-05 - Created script

.PARAMETER ComputerName
    One or more computer names

.PARAMETER Credential
    Valid credentials on target computer (default is current user credentials)

.PARAMETER AsJob
    Invoke the command as a PSjob

.PARAMETER ConnectionTest
    Run WinRM or ping test prior to invoke-command, or no test (default is WinRM)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> 

        
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
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Hostname or FQDN of computer (separate multiple computers with commas)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

        [Parameter(
        HelpMessage = "Return network and disk data as collections or as a summary"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Collections","Summary")]
    [string]$OutputType = "Summary",

    [Parameter(
        Mandatory=$False,
        HelpMessage="Valid credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Run Invoke-Command scriptblock as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Test to run prior to invoke-command - WinRM (default, using Kerberos), ping, or none)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)
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
    
    # Output
    [array]$Results = @()
    
    #region Scriptblock for invoke-command

    $ScriptBlock = {
        
        Param($OutputType)

        $ErrorActionPreference = "Stop"
        
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
            "Messages"
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
 
        #region Inner functions

        # Function to get pending reboot
        Function Test-PendingReboot{        
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
        $Activity = "Querying computer"
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
            Class        = $Null
            ErrorAction  = "Stop"
            Verbose      = $False
        } 
    
        #endregion Splats

        Try  { 
            
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
                $RebootPending = Test-PendingReboot
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
            $Forward = Test-DNS -Name $Env:ComputerName
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

            $TotalGB = [math]::Round(($ObjDisk.SizeGB | Measure-Object -Sum).Sum)
            $TotalFree = [math]::Round(($ObjDisk.FreeGB | Measure-Object -Sum).Sum)

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
                TotalGB             = $TotalGB
                TotalFree           = $TotalFree
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
                Messages            = $Null 
            }  | Select $Select_CompInfo

            

            Switch ($Using:OutputType) {
                Collections {
                    $ObjComputer.HardDriveInfo = $ObjDisk
                    $ObjComputer.NetAdapterInfo = $ObjNetwork
                }               
                Summary {
                    $DiskSummary = ($ObjDisk | ForEach-Object {"$($_.DeviceID) [$($_.VolumeName)] $($_.SizeGB)GB ($($_.PercentFree)% free)"}) -join(", ")
                    $ObjComputer.HardDriveInfo = $DiskSummary
                    
                    $NetworkSummary = ($ObjNetwork | 
                        ForEach-Object {
                            "[$($_.NetConnectionID)] MACAddress: $($_.MACAddress); IPv4address: $(($_.IPAddress | Where-Object {$_ -notmatch ":"}) -join(", "))/$Prefix, Gateway: $($_.DefaultGateway); DNS: $($_.DNSServerOrder -join(", ")); DHCPEnabled: $($_.DHCPEnabled) "}
                    ) -join(", ")
                    $ObjComputer.NetAdapterInfo = $NetworkSummary
                }
            }  
            
            #endregion Compinfo object
            
        }
        Catch { 
            
            $ObjComputer = New-Object PSObject -Property @{ 
                ComputerName        = $Env:ComputerName
                Domain              = $Null
                OperatingSystem     = $Null
                OSArchitecture      = $Null
                BuildNumber         = $Null
                ServicePack         = $Null
                IsVirtual           = $Null
                Manufacturer        = $Null
                Model               = $Null
                SerialNumber        = $Null
                Processor           = $Null
                LogicalProcessor    = $Null
                PhysicalMemoryGB    = $Null
                OSReportedMemoryGB  = $Null
                HardDriveCount      = $Null
                HardDriveInfo       = $Null
                TotalGB             = $Null
                TotalFree           = $Null
                NetAdapterCount     = $Null
                NetAdapterInfo      = $Null
                MACAddress          = $Null
                IPAddress           = $Null
                PAEEnabled          = $Null
                ForwardDNSLookup    = $Null
                ReverseDNSLookup    = $Null
                InstallDate         = $Null
                LastBootUpTime      = $Null
                UpTime              = $Null
                RebootPending       = $Null
                Messages            = $_.Exception.Message 
            }  | Select $Select_CompInfo
        }

        Write-Output $ObjComputer

    } #end scriptblock

    #endregion Scriptblock for invoke-command

    #region Functions

    Function Test-WinRM{
        Param($Computer)
        $Param_WSMAN = @{
            ComputerName   = $Computer
            Credential     = $Credential
            Authentication = "Kerberos"
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
    $Activity = "Get Windows report data"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as remote PSJob"
    }
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        ID               = 1
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    # Parameters for Invoke-Command
    $ConfirmMsg = $Activity
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName   = ""
        Authentication = "Kerberos"
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.AsJob = $True
        $Param_IC.JobName = $Null
        $JobPrefix = "WinReport"
    }
    
    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "$Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}


} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        $Computer = $Computer.Trim()
        $Current ++ 

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = $Computer
        
        [switch]$Continue = $False

        Switch ($ConnectionTest) {
            Default {$Continue = $True}
            Ping {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Ping computer"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "Ping failure"
                            $Host.UI.WriteErrorLine("[$Computer] $Msg")
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        $Host.UI.WriteErrorLine("[$Computer] $Msg")
                    }
                }
                Else {$Continue = $True}
            }
            WinRM {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Test WinRM connection"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-WinRM -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "WinRM failure"
                            $Host.UI.WriteErrorLine("[$Computer] $Msg")
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        $Host.UI.WriteErrorLine("[$Computer] $Msg")
                    }
                }
                Else {$Continue = $True}
            }        
        }

        If ($Continue.IsPresent) {
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Msg = "Invoke command"
                    If ($AsJob.IsPresent) {$Msg += " as PSJob"}
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Jobs += $Job
                    }
                    Else {
                        $Results += Invoke-Command @Param_IC
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                    $Host.UI.WriteErrorLine("[$Computer] $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                $Host.UI.WriteErrorLine("[$Computer] $Msg")
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {

        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            Write-Verbose $Msg
            $Jobs | Get-Job
        }
        Else {
            $Msg = "No jobs created"
            Write-Warning $Msg
        }
    } #end if AsJob

    Else {
        If ($Results.Count -eq 0) {
            $Msg = "No results found"
            Write-Warning $Msg
        }
        Else {
            $Msg = "$($Results.Count) result(s) found"
            Write-Verbose $Msg
            Write-Output ($Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID)
        }
    }

}

} # end Do-SomethingCool

