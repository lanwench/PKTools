#Requires -version 4
Function Get-PKvCenterMetrics {
<# 
.SYNOPSIS
    Lists all available vCenter performance metrics, optionally sorted or grouped by property names

.DESCRIPTION
    Lists all available vCenter performance metrics, optionally sorted or grouped by property names
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Get-PKvCenterMetrics.ps1
    Created : 2019-07-09
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-07-09 - Created script

.LINK
    https://orchestration.io/2015/05/18/listing-all-vcenter-performance-metrics-with-powercli/

.LINK
    https://www.vmware.com/support/developer/converter-sdk/conv43_apireference/vim.ServiceInstance.html

.PARAMETER VIServer
    vCenter server name or object (default is any connected)

.PARAMETER SortBy
    Property to sort by: GroupKey, NameKey, RollupType, Level, FullName (if null, is Default)

.PARAMETER GroupBy
    Property to group by: GroupKey, NameKey, RollupType, Level

.PARAMETER Quiet
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKvCenterMetrics -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                  
        ---              -----                  
        Verbose          True                   
        VIServer                                
        SortBy                                  
        Quiet            False                  
        PipelineInput    False                  
        ParameterSetName __DefaultParameterSet  
        ScriptName       Get-PKvCenterMetrics
        ScriptVersion    1.0.0                  

        VERBOSE: [Prerequisites] Verify PowerCLI module is available
        VERBOSE: [Prerequisites] Verified PowerCLI module 11.0.0 in C:\Program Files\WindowsPowerShell\Modules\VMware.VimAutomation
        .Core\11.0.0.10336080
        VERBOSE: [Prerequisites] Test vCenter connectivity
        VERBOSE: [Prerequisites] Connection found to vCenter 'https://vcenter.domain.local/sdk,' version 6.5.0

        BEGIN: List all available vCenter performance metrics

        [vcenter.domain.local] Get ServiceInstance view
        [vcenter.domain.local] Get performance counters

        GroupKey   : cpu
        NameKey    : usage
        RollupType : none
        Level      : 4
        FullName   : cpu.usage.none
        Summary    : CPU usage as a percentage during the interval

        GroupKey   : cpu
        NameKey    : usage
        RollupType : average
        Level      : 1
        FullName   : cpu.usage.average
        Summary    : CPU usage as a percentage during the interval

        GroupKey   : cpu
        NameKey    : usage
        RollupType : minimum
        Level      : 4
        FullName   : cpu.usage.minimum
        Summary    : CPU usage as a percentage during the interval

        GroupKey   : cpu
        NameKey    : usage
        RollupType : maximum
        Level      : 4
        FullName   : cpu.usage.maximum
        Summary    : CPU usage as a percentage during the interval

        GroupKey   : cpu
        NameKey    : usagemhz
        RollupType : none
        Level      : 4
        FullName   : cpu.usagemhz.none
        Summary    : CPU usage in megahertz during the interval


.EXAMPLE
    PS C:\> Get-PKvCenterMetrics -SortBy FullName -Quiet | Format-Table -Autosize

        VERBOSE: PSBoundParameters: 
	
        Key              Value                  
        ---              -----                  
        Verbose          True                   
        SortBy           RollupType             
        VIServer                                
        Quiet            False                  
        PipelineInput    False                  
        ParameterSetName __DefaultParameterSet  
        ScriptName       Get-PKvCenterMetrics
        ScriptVersion    1.0.0                  

        VERBOSE: [Prerequisites] Verify PowerCLI module is available
        VERBOSE: [Prerequisites] Verified PowerCLI module 11.0.0 in C:\Program Files\WindowsPowerShell\Modules\VMware.VimAutomation.Core\11.0.0.10336080
        VERBOSE: [Prerequisites] Test vCenter connectivity
        VERBOSE: [Prerequisites] Connection found to vCenter 'vcenter.domain.local,' version 6.5.0

        BEGIN: List all available vCenter performance metrics, sorted by RollupType

        [vcenter.domain.local] Get ServiceInstance view from vCenter https://vcenter.domain.local/sdk
        [vcenter.domain.local] Get performance counters

        GroupKey        NameKey                        RollupType Level FullName                                 Summary                                                                                                       
        --------        -------                        ---------- ----- --------                                 -------                                                                                                       
        mem             latency                           average     2 mem.latency.average                      Percentage of time the virtual machine spent waiting to swap in or decompress guest physical memory           
        managementAgent memUsed                           average     3 managementAgent.memUsed.average          Amount of total configured memory that is available for use                                                   
        managementAgent swapUsed                          average     3 managementAgent.swapUsed.average         Sum of the memory swapped by all powered-on virtual machines on the host                                      
        power           capacity.usage                    average     4 power.capacity.usage.average             Current power usage                                                                                           
        power           capacity.usable                   average     4 power.capacity.usable.average            Current maximum allowed power usage.                                                                          
        net             throughput.usage.hbr              average     3 net.throughput.usage.hbr.average         Average pNic I/O rate for HBR                                                                                 
        net             throughput.usage.iscsi            average     3 net.throughput.usage.iscsi.average       Average pNic I/O rate for iSCSI                                                                               
        net             throughput.usage.ft               average     3 net.throughput.usage.ft.average          Average pNic I/O rate for FT                                                                                  
        net             throughput.usage.vmotion          average     3 net.throughput.usage.vmotion.average     Average pNic I/O rate for vMotion                                                                             
        net             throughput.usage.nfs              average     3 net.throughput.usage.nfs.average         Average pNic I/O rate for NFS                                                                                 
        net             throughput.usage.vm               average     3 net.throughput.usage.vm.average          Average pNic I/O rate for VMs                                                                                 
        mem             activewrite                       average     2 mem.activewrite.average                  Amount of guest physical memory that is being actively written by guest. Activeness is estimated by ESXi      
        managementAgent cpuUsage                          average     3 managementAgent.cpuUsage.average         Amount of Service Console CPU usage                                                                           
        storagePath     commandsAveraged                  average     3 storagePath.commandsAveraged.average     Average number of commands issued per second on the storage path during the collection interval               
        storageAdapter  throughput.usag                   average     4 storageAdapter.throughput.usag.average   The storage adapter's I/O rate                                                                                
        storageAdapter  queueLatency                      average     2 storageAdapter.queueLatency.average      Average amount of time spent in the VMkernel queue, per SCSI command, during the collection interval         

        <snip>

.EXAMPLE
    PS C:\> Get-PKvCenterMetrics -GroupBy GroupKey -OutVariable Metrics

        Name               Value                                                                                      
        ----               -----                                                                                      
        vcResources        {@{GroupKey=vcResources; NameKey=buffersz; RollupType=average; Level=4; FullName=vcResou...
        vmop               {@{GroupKey=vmop; NameKey=numPoweron; RollupType=latest; Level=1; FullName=vmop.numPower...
        mem                {@{GroupKey=mem; NameKey=usage; RollupType=none; Level=4; FullName=mem.usage.none; Summa...
        storageAdapter     {@{GroupKey=storageAdapter; NameKey=commandsAveraged; RollupType=average; Level=2; FullN...
        vcDebugInfo        {@{GroupKey=vcDebugInfo; NameKey=activationlatencystats; RollupType=maximum; Level=4; Fu...
        datastore          {@{GroupKey=datastore; NameKey=numberReadAveraged; RollupType=average; Level=1; FullName...
        gpu                {@{GroupKey=gpu; NameKey=utilization; RollupType=none; Level=4; FullName=gpu.utilization...
        net                {@{GroupKey=net; NameKey=usage; RollupType=none; Level=4; FullName=net.usage.none; Summa...
        managementAgent    {@{GroupKey=managementAgent; NameKey=swapOut; RollupType=average; Level=3; FullName=mana...
        virtualDisk        {@{GroupKey=virtualDisk; NameKey=numberReadAveraged; RollupType=average; Level=1; FullNa...
        pmem               {@{GroupKey=pmem; NameKey=available.reservation; RollupType=latest; Level=4; FullName=pm...
        vsanDomObj         {@{GroupKey=vsanDomObj; NameKey=readIops; RollupType=average; Level=4; FullName=vsanDomO...
        power              {@{GroupKey=power; NameKey=power; RollupType=average; Level=2; FullName=power.power.aver...
        vflashModule       {@{GroupKey=vflashModule; NameKey=numActiveVMDKs; RollupType=latest; Level=4; FullName=v...
        sys                {@{GroupKey=sys; NameKey=uptime; RollupType=latest; Level=1; FullName=sys.uptime.latest;...
        rescpu             {@{GroupKey=rescpu; NameKey=actav1; RollupType=latest; Level=3; FullName=rescpu.actav1.l...
        clusterServices    {@{GroupKey=clusterServices; NameKey=cpufairness; RollupType=latest; Level=1; FullName=c...
        disk               {@{GroupKey=disk; NameKey=usage; RollupType=none; Level=4; FullName=disk.usage.none; Sum...
        hbr                {@{GroupKey=hbr; NameKey=hbrNumVms; RollupType=average; Level=4; FullName=hbr.hbrNumVms....
        cpu                {@{GroupKey=cpu; NameKey=usage; RollupType=none; Level=4; FullName=cpu.usage.none; Summa...
        storagePath        {@{GroupKey=storagePath; NameKey=throughput.cont; RollupType=average; Level=4; FullName=...


    PS C:\> $Metrics.power | Format-Table -AutoSize

        GroupKey NameKey           RollupType Level FullName                        Summary                                       
        -------- -------           ---------- ----- --------                        -------                                       
        power    power                average     2 power.power.average             Current power usage                           
        power    powerCap             average     3 power.powerCap.average          Maximum allowed power usage                   
        power    energy             summation     3 power.energy.summation          Total energy used since last stats reset      
        power    capacity.usagePct    average     4 power.capacity.usagePct.average Current power usage as a percentage of maxi...
        power    capacity.usable      average     4 power.capacity.usable.average   Current maximum allowed power usage.          
        power    capacity.usage       average     4 power.capacity.usage.average    Current power usage                           


#> 

[CmdletBinding(
    DefaultParameterSetName = "Sort",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage = "vCenter server name or object (default is any connected)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name")]
    [object]$VIServer,


    [Parameter(
        ParameterSetName = "Sort",
        HelpMessage = "Property to sort by: GroupKey, NameKey, RollupType, Level, FullName (if null, is Default)"
    )]
    [ValidateSet("GroupKey","NameKey","RollupType","Level","FullName")]
    [string]$SortBy,

    [Parameter(
        ParameterSetName = "Group",
        Mandatory = $True,
        HelpMessage = "Property to group by: GroupKey, NameKey, RollupType, Level"
    )]
    [ValidateSet("GroupKey","NameKey","RollupType","Level")]
    [string]$GroupBy,

    [Parameter(
        HelpMessage = "Suppress non-verbose console output"
    )]
    [Switch]$Quiet

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
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
    
    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    #region Functions

    # Function to test vCenter connection
    Function TestVISession {
    [CmdletBinding()]
    Param([string]$Server,[switch]$Bool)
        $VIServer = $Null
        If ($PSBoundParameters.Server) {
            Try { 
                If ($ServerObj = Get-Variable DefaultVIServers -Scope Global -ErrorAction SilentlyContinue | Select -ExpandProperty Value | Where-Object {$_.Name -match $Server} ) {
                    If ($Bool.IsPresent) {$True}
                    Else {Write-Output $ServerObj}
                }
                Else {
                    $Msg = "No connection found to vCenter server '$Server'"
                    If ($Bool.IsPresent) {$False}   
                }
            }Catch {}
        }
        Else {
            Try {
                If ($ServerObj = Get-Variable DefaultVIServers -Scope Global -ErrorAction SilentlyContinue | Select -ExpandProperty Value) {
                    If ($Bool.IsPresent) {$True}
                    Else {
                        Write-Output $ServerObj
                    }
                }
                Else {
                    $Msg = "No connection found to vCenter"
                    If ($Bool.IsPresent) {$False}
                }
            }Catch {}
        }
    } #end TestViSession

    # Function to write a console message or a verbose message
    Function Write-MessageInfo {
        Param([Parameter(ValueFromPipeline)]$Message,$FGColor,[switch]$Title)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {
            If ($Title.IsPresent) {$Message = "`n$Message`n"}
            $Host.UI.WriteLine($FGColor,$BGColor,"$Message")
        }
        Else {Write-Verbose "$Message"}
    } #end Write-MessageInfo

    # Function to write an error or a verbose message
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)#,[switch]$Quiet = $Quiet)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("$Message")}
        Else {Write-Verbose "$Message"}
    } #end Write-MessageError

    #endregion Functions
    
    #region Prerequisites
    
    $Activity = "Prerequisites"

    $Msg = "Verify PowerCLI module is available"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Activity -CurrentOperation $Msg
    Try {
        If (-not ($PCLIMod = Get-Module VMware.vimautomation.core -ErrorAction SilentlyContinue -Verbose:$False -Debug:$False)) {
            $Msg = "Failed to detect VMware PowerCLI module loaded in this session; it can be installed from https://www.powershellgallery.com/packages/VMware.PowerCLI"
            $Host.UI.WriteErrorLine("[Prerequisites] $Msg")
            Break
        }
        Else {
            $Msg = "Verified PowerCLI module $($PCLIMod.Version.Major).$($PCLIMod.Version.Minor).$($PCLIMod.Version.Build) in $($PCLIMod.Path | Split-Path -Parent)"
            Write-Verbose "[Prerequisites] $Msg"
        }
    }
    Catch {
        $Msg = "Failed to detect VMware PowerCLI module loaded in this session; please ensure it is installed from https://www.powershellgallery.com/packages/VMware.PowerCLI"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
        $Host.UI.WriteErrorLine("[Prerequisites] $Msg")
        Break
    }

    $Msg = "Test vCenter connectivity"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    If ($CurrentParams.VIServer) {
        If (TestVISession -Server $VIServer -Verbose:$False) {
            $Msg = "Connection found to '$($VIServer.ServiceUri),' version $($VIServer.Version)"
            Write-Verbose "[Prerequisites] $Msg"
        }
        Else {
            $Msg = "No connection found to vCenter '$VIServer'"
            "[Prerequisites] $Msg" | Write-MessageError
            Break
        }
    }
    Else {
        If ($VIServer = TestVISession -Verbose:$False) {
            If ($VIServer.Count -gt 1) {
                $Msg = "Multiple vCenter connections found: $($VIServer.Name -join(', '))`nPlease re-run function with the -VIServer parameter"
                "[Prerequisites] $Msg" | Write-MessageError
                Break
            }
            Else {
                $Msg = "Connection found to '$($VIServer.ServiceUri),' version $($VIServer.Version)"
                Write-Verbose "[Prerequisites] $Msg"
            }
        }
        Else {
            $Msg = "No connection found to vCenter"
            "[Prerequisites] $Msg" | Write-MessageError
            Break
        }
    }
    
    #endregion Prerequisites
   
    #region Splats

    # General-purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
        Debug       = $False
    }

    # For Write-Progress
    $Activity = "List all available vCenter performance metrics"
    Switch ($Source) {
        Sort {If ($CurrentParams.SortBy) {$Activity += " (sorted by $SortBy)"}}
        Group {$Activity += " (grouped by $GroupBy)"}
    }
    
    #endregion Splats
 
    # Console output
    $Msg = "BEGIN: $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title


} #end begin

Process {

    $Msg = "Get ServiceInstance view"
    Write-MessageInfo "[$($VIServer.Name)] $Msg" -FGColor White

    Try {
        
        # define vCenter service instance and performance manager
        $serviceInstance = Get-View ServiceInstance -Server $VIServer @StdParams
        $perfMgr = Get-View $serviceInstance.Content.PerfManager -Server $VIServer @StdParams 
        
        # get all available performance counters & output a PSObject with details
        $counters = $perfMgr.PerfCounter
        $Output = @()
        $Msg = "Get performance counters"
        Write-MessageInfo "[$($VIServer.Name)] $Msg" -FGColor White
        
        $Total = $Counters.Count
        $Current = 0
        
        Foreach ($counter in $counters) {
            $FullName = "$($counter.GroupInfo.Key).$($counter.NameInfo.Key).$($counter.RollupType)"
            $Current ++
            Write-Progress -Activity $Activity -CurrentOperation $FullName -PercentComplete ($Current/$Total*100)
            $Output += [PSCustomObject]@{
                GroupKey   = $Counter.GroupInfo.Key
                NameKey    = $counter.NameInfo.Key
                RollupType = $counter.RollupType
                Level      = $counter.Level
                FullName   = $FullName
                Summary    = $counter.NameInfo.Summary
            }
        }

        Write-Progress -Activity $Activity -CurrentOperation $Null
        
        Switch ($Source) {
            Sort {
                If ($CurrentParams.SortBy) {
                    Write-Output ($Output | Sort-Object $SortBy)
                }
                Else {
                    Write-Output $Output
                }
            }
            Group {
                Write-Output ($Output | Group-Object -Property $GroupBy -AsHashtable -AsString)
            }
        }
    }
    Catch {
        $Msg = "Failed to get ServiceInstance view"
        If ($ErrorDetails = $($_.exception.message -replace '\s+', ' ')) {$Msg += "`n$ErrorDetails"}
        "[$($VIServer.Name)] $Msg" | Write-MessageError 
    }
  
}
End {
    
    # Console output
    $Msg = "END  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title
    $Null = Write-Progress -Activity $Activity -Completed

}

} # end Get-PKvCenterMetrics

