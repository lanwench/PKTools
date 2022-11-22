#requires -version 4
function Get-PKSID {
<#
.SYNOPSIS
    Gets the SID for one or more local or domain users or groups via .NET

.DESCRIPTION
    Gets the SID for one or more local or domain users or groups via .NET
    Uses System.Security.Principal.NTAccount to translate the identity into a SID
    Defaults to current user
    Accepts pipeline input
    Returns a PSObject

.NOTES        
    Name    : Function_Get-PKSID.ps1
    Author  : Paula Kingsley
    Created : 2022-10-20
    Version : 01.00.000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2022-10-20 - Created script

.LINK
    https://unlockpowershell.wordpress.com/2009/11/20/script-remote-dcom-wmi-access-for-a-MEGACORP-user/

.PARAMETER Identity
    One or more local or domain user or group names (default is current user)

.EXAMPLE
    PS C:\> Get-PKSID -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value         
        ---           -----         
        Verbose       True          
        Identity      {mburns}
        ScriptName    Get-PKSID     
        PipelineInput False         
        ScriptVersion 1.0.0         

        VERBOSE: [mburns] Translating name to SID

        Name      SID                                         
        ----      ---                                         
        mburns    S-1-5-21-2503949928-964973733-913657002-1549


.EXAMPLE
    PS C:\> Get-PKSID mburns,MEGACORP\jbloggs,ACQUISITION\ephipps,MEGACORP\helpdesk_admins 

        WARNING: [ACQUISITION\ephipps] Exception calling "Translate" with "1" argument(s): "Some or all identity references could not be translated."
    
        Name                      SID                                                                                                           
        ----                      ---                                                                                                           
        mburns                    S-1-5-21-2503949928-964973733-913657002-1549
        MEGACORP\jbloggs          S-1-5-21-2523766570-455092341-570019136-24612                                                                 
        ACQUISITION\ephipps       Exception calling "Translate" with "1" argument(s): "Some or all identity references could not be translated."
        MEGACORP\helpdesk_admins  S-1-5-21-2523766570-455092341-570019136-19854             
    
.EXAMPLE
    PS C:\> Get-LocalUser | Get-PKSid 

        Name               SID                                           
        ----               ---                                           
        Administrator      S-1-5-21-3715883451-8334022814-3290647066-500 
        DefaultAccount     S-1-5-21-3715883451-8334022814-3290647066-503 
        Guest              S-1-5-21-3715883451-8334022814-3290647066-501 
        jack               S-1-5-21-3715883451-8334022814-3290647066-1001
        WDAGUtilityAccount S-1-5-21-3715883451-8334022814-3290647066-504     

#>

[CmdletBinding()]
Param (
    [Parameter(
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        Position = 0,
        HelpMessage = "One or more local or domain user or group names (default is current user)"  
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name")]
    [object[]]$Identity = $env:USERNAME
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Activity = "Get SID for users or groups"
    Write-Verbose "[BEGIN: $ScCriptname] $Activity"

}
Process {

    $Total = $Identity.Count
    $Current = 0
    Foreach ($Name in $Identity) {
        $Current ++
        Try {
            $Msg = "Translating name to SID"
            Write-Verbose "[$Name] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Name -Status $Msg -PercentComplete ($Current/$Total*100)
            $ID = new-object System.Security.Principal.NTAccount($Name)
            $SID = $ID.Translate( [System.Security.Principal.SecurityIdentifier] ).toString()
            
            [PSCustomObject]@{
                Name = $Name
                SID  = $SID
            }
        }
        Catch {
            Write-Warning "[$Name] $($_.Exception.Message)"
            
            [PSCustomObject]@{
                Name = $Name
                SID = $_.Exception.Message
            }
        }
    }
}
End {
    Write-Verbose "[END: $ScCriptname] $Activity"
}
} #end Get-PKSID

    