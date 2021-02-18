#requires -Version 4
Function Get-PKVMIPConfig {
<#
.SYNOPSIS
    Returns IPv4 configuration data for one or more guest VMs

.DESCRIPTION
    Returns IPv4 configuration data for one or more guest VMs
    Accepts strings or VM objects
    Uses Get-View
    Accepts pipeline input
    Returns a PSObject

.NOTES
    Name    : Function_Get-PKVMIPconfig.ps1
    Author  : Paula Kingsley
    Created : 2020-12-27
    Version : v1.00.0000
    History : 

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2020-12-27 - Created script based on Luk Dekens' original

.LINK
    https://communities.vmware.com/t5/VMware-PowerCLI-Discussions/PowerCLI-Script-to-Gather-VM-Guest-Network-Information-Using-get/td-p/510682

.PARAMETER VM
    One or more VM names or objects

.PARAMETER VIServer
    One or more vCenter servers (default is any connected)

.EXAMPLE
    PS C:\> Get-PKVMIPConfig -VM sql2 -VIServer vcenter.domain.local -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                              
        ---           -----                              
        VM            sql2                       
        VIServer      vcenter.domain.local
        Verbose       True                               
        ScriptName    Get-PKVMIPConfig                   
        PipelineInput False                              
        ScriptVersion 1.0.0                              

        VERBOSE: [sql2] Find VM
        VERBOSE: [sql2] Look up guest IP config


        Name           : sql2
        IPAddress      : 10.19.74.128
        Gateway        : 10.19.74.1
        SubnetMask     : 255.255.255.0
        DNSHostname    : sql2.domain.local
        DNS            : {10.20.2.14}
        MACAddress     : 00:50:82:96:ce:e2
        Connected      : True
        StartConnected : True
        VLANId         : infra
        NetworkAdapter : Vmxnet3
        VIServer       : vcenter.domain.local


.EXAMPLE
    PS C:\>  Get-VM *rdp*,devbox123 | get-PKVMIPConfig | Format-Table -AutoSize

        Name          IPAddress     Gateway     SubnetMask    DNSHostname                DNS                         MACAddress        Connected StartConnected VLANId                    
        ----          ---------     -------     ----------    -----------                ---                         ----------        --------- -------------- ------                    
        websrv-rdp-1  10.10.32.332  10.10.32.1  255.255.254.0 websrv-rdp-1.domain.local  {10.10.62.30, 192.158.30.3} 00:50:56:93:a5:14      True           True prod
        supportrdp    10.10.33.19   10.10.33.1  255.255.254.0 supportrdp.domain.local    {10.10.62.30, 192.158.30.3} 00:50:56:93:60:57      True           True prod
        dardp         10.10.33.20   10.10.33.1  255.255.254.0 dardrdp.domain.local       {10.10.62.30, 192.158.30.3} 00:50:56:93:2a:eb      True           True prod
        devbox123     172.30.52.49  172.30.52.1 255.255.255.0 devbox123.acquisition.co   {172.30.67.1, 8.8.8.8}      00:50:56:96:67:81      True           True 12_dev


#>
[Cmdletbinding()]
Param(
    [Parameter(
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelinebyPropertyName = $True,
        HelpMessage = "One or more VM names or objects"
    )]
    [Alias("Name","VMName")]
    [ValidateNotNullOrEmpty()]
    $VM,

    [Parameter(
        HelpMessage = "One or more vCenter servers (default is any connected)"
    )]
    [Alias("Server")]
    [ValidateNotNullOrEmpty()]
    $VIServer = $Global:DefaultVIServers
)

Begin{

    # Current version (please keep up to date from comment block)
    $Version = [version]"01.00.0000"

    # How did we get here?
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $ScriptName = $MyInvocation.MyCommand.Name
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Activity = "Get VM guest IPv4 configuration"

    If (-not ($Null = Get-Module VMware.PowerCLI)) {
        Write-Error "PowerCLI module not found!"
        Break
    }

}
Process {
    
    $Total = ($VM -as [array]).count
    $Current = 0

    Foreach ($V in $VM) {
        
        $Current ++    
        
        # If it's a string
        If (-not ($V -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl])) {

            $Msg = "[$V] Find VM"
            Write-Verbose $Msg
            Write-Progress -Activity $Activity -CurrentOperation $Msg -PercentComplete ($Current/$Total*100) 
            Try {
                $VMObj = Get-VM $V -Server $VIServer -Verbose:$False -ErrorAction Stop
            }
            Catch {
                $Msg = "[$V] $($_.exception.message -replace '\s+', ' ')"
                Write-Error $Msg
            }
        }
        Elseif ($V -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {
            $VMObj = $V
        }
        
        If ($VMObj) {
                
            $Msg = "[$($VMObj.Name)] Look up guest IP config"
            Write-Verbose $Msg
            Write-Progress -Activity $Activity -CurrentOperation $Msg -PercentComplete ($Current/$Total*100) 
                
            Try {
                Get-view -viewtype VirtualMachine -filter @{Name="^$($VMObj.Name)$"} -PipelineVariable VM -Server $VIServer -Verbose:$False -ErrorAction Stop |
                    Select Name,
                    @{N='IPAddress';E={$_.Guest.ipAddress}},
                    @{N='Gateway';E={($_.Guest.ipstack.iprouteconfig.iproute.gateway.ipaddress | where{$_ -ne $null})}},
                    @{N='SubnetMask';E={
                        $IPAddr = $_.Guest.ipAddress
                        @(($_.Guest.Net.ipconfig.ipaddress | Where-Object {$IPAddr -contains $_.IpAddress -and $_.PrefixLength -ne 0}).PrefixLength | Foreach-Object {
                            [IPAddress]$ip = 0;
                            $ip.Address = (([UInt32]::MaxValue) -shl (32 - $_) -shr (32 - $_))
                            $ip.IPAddressToString
                        })
                    }},
                    @{N='DNSHostname';E={$_.Guest.Hostname}},
                    @{N='DNS';E={$_.Guest.IpStack.DnsConfig.IpAddress}},
                    @{N='MACAddress';E={($_.Config.Hardware.Device | Where-Object {$_ -is [VMware.Vim.VirtualEthernetCard]}).MacAddress}},
                    @{N='Connected';E={
                        ($_.Config.Hardware.Device | Where-Object {$_ -is [VMware.Vim.VirtualEthernetCard]}).Connectable.Connected
                    }},
                    @{N='StartConnected';E={
                        ($_.Config.Hardware.Device | Where-Object {$_ -is [VMware.Vim.VirtualEthernetCard]}).Connectable.StartConnected
                    }},
                    @{N='VLANId';E={
                        $folder = Get-View -Id $_.Parent -Property "[VirtualMachine]ChildEntity.Network.*" -Server $VIServer -Verbose:$False
                        ($folder.LinkedView.ChildEntity.where({$vm.MoRef -eq $_.MoRef})).LinkedView.Network.Name
                    }},
                    @{N='NetworkAdapter';E={$_.Config.Hardware.Device  |  
                        Where-Object {$_ -is [VMware.Vim.VirtualEthernetCard]} | 
                            ForEach-Object {$_.GetType().Name.Replace('Virtual','')}
                    }},
                    @{N='VIServer';E={([system.uri]$VM.Client.ServiceURL).Host}}
                }
                Catch {
                    $Msg = "[$($VMObj.Name)] $($_.exception.message -replace '\s+', ' ')"
                    Write-Error $Msg
                }
        } # end if VM object
        
    } #end foreach

        
} #end process

} #end Get-PKVMIPconfig
