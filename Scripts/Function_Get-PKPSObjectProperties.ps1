#Requires -version 3
Function Get-PKPSObjectProperties {
<# 
.SYNOPSIS
    Returns property names from a PSCustomObject in order as an array, suitable for Select-Object

.DESCRIPTION
    Returns property names from a PSCustomObject in order as an array, suitable for Select-Object
    Accepts pipeline input
    Returns an array of strings

.NOTES        
    Name    : Function_Get-PKPSObjectProperties.ps1
    Created : 2018-10-23
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-10-23 - Created script

.PARAMETER PSObject
    PSCustomObject to retrieve property names from

.EXAMPLE
    PS C:\> Get-WMIObject -Class Win32_ComputerSystem | Get-PKPSObjectProperties -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                   
        ---           -----                   
        Verbose       True                    
        PSObject                              
        ScriptName    Get-PKPSObjectProperties
        ScriptVersion 1.0.0                   

        VERBOSE: Return array of property names in order (e.g., useful for Select-Object in downlevel PS versions not supporting ordered hashtables)

        'PSComputerName','__GENUS','__CLASS','__SUPERCLASS','__DYNASTY','__RELPATH','__PROPERTY_COUNT','__DERIVATION','__SERVER','__NAMESPACE','__PATH
        ','AdminPasswordStatus','AutomaticManagedPagefile','AutomaticResetBootOption','AutomaticResetCapability','BootOptionOnLimit','BootOptionOnWatc
        hDog','BootROMSupported','BootStatus','BootupState','Caption','ChassisBootupState','ChassisSKUNumber','CreationClassName','CurrentTimeZone','D
        aylightInEffect','Description','DNSHostName','Domain','DomainRole','EnableDaylightSavingsTime','FrontPanelResetStatus','HypervisorPresent','In
        fraredSupported','InitialLoadInfo','InstallDate','KeyboardPasswordStatus','LastLoadInfo','Manufacturer','Model','Name','NameFormat','NetworkSe
        rverModeEnabled','NumberOfLogicalProcessors','NumberOfProcessors','OEMLogoBitmap','OEMStringArray','PartOfDomain','PauseAfterReset','PCSystemT
        ype','PCSystemTypeEx','PowerManagementCapabilities','PowerManagementSupported','PowerOnPasswordStatus','PowerState','PowerSupplyState','Primar
        yOwnerContact','PrimaryOwnerName','ResetCapability','ResetCount','ResetLimit','Roles','Status','SupportContactDescription','SystemFamily','Sys
        temSKUNumber','SystemStartupDelay','SystemStartupOptions','SystemStartupSetting','SystemType','ThermalState','TotalPhysicalMemory','UserName',
        'WakeUpType','Workgroup','Scope','Path','Options','ClassPath','Properties','SystemProperties','Qualifiers','Site','Container'

.EXAMPLE
    PS C:\> Get-PKPSObjectProperties -PSObject $Output

        'ComputerName','NetConnectionID','Description','Index','MACAddress','IsDHCP','IPv4Address','SubnetMask','Gateway',
        'DNSServerSearchOrder','DNSDomainSearchOrder','DomainDNSRegistrationEnabled','DNSDomain','DHCPServer',
        'DHCPLeaseObtained','Messages'
        
#> 

[CmdletBinding()]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="PSCustomObject"
    )]
    [ValidateNotNullOrEmpty()]
    [pscustomobject] $PSObject
)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
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
   
} #end begin

Process {

    Try {
        $Msg = "Return array of property names (e.g., useful for Select-Object in downlevel PS versions not supporting ordered hashtables)"
        Write-Verbose $Msg
        "'$($PSObject.PSObject.Properties.Name -join("','"))'"
    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $Msg"}
        $Host.UI.WriteErrorLine($Msg)
    }
   
}
End {}

} # end Get-PKPSObjectProperties

