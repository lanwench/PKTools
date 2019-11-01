#requires -Version 3
Function Get-PKCurrentUser {
<#
.SYNOPSIS 
    Gets details for the currently logged-in user, including group membership, using .NET DirectoryServices.AccountManagement

.DESCRIPTION
    Gets details for the currently logged-in user, including group membership, using .NET DirectoryServices.AccountManagement
    No ActiveDirectory module required
    Optional parameters look up group SIDs to names, and expand group name collections to stacked strings
    Outputs a PSObject
    
.OUTPUTS
    PSObject

.NOTES        
    Name    : Function_Get-PKCurrentUser.ps1 
    Author  : Paula Kingsley
    Created : 2019-10-31
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
    
        v01.00.0000 - 2019-10-31 - Created script based on link

.PARAMETER LookupGroupNames
    Attempt lookup of group SIDs to names

.PARAMETER ExpandGroupNames
    Expand group names in collection, to stacked strings (presumes -LookupGroupNames)

.PARAMETER Quiet
    Suppress non-verbose console output

.LINK
    https://powershellposse.com/2019/04/10/user-group-information-directoryservices-accountmanagement/

.EXAMPLE
    PS C:\> Get-PKCurrentUser -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value            
        ---              -----            
        Verbose          True             
        LookupGroupNames False            
        ExpandGroupNames False            
        Quiet            False            
        ParameterSetName Default          
        ScriptName       Get-PKCurrentUser
        ScriptVersion    1.0.0            

        BEGIN: Look up current user

        ComputerName       : LAPTOP
        AuthenticationType : Kerberos
        ImpersonationLevel : None
        Name               : DOMAIN\jbloggs
        IsAuthenticated    : True
        IsSystem           : False
        SID                : S-1-5-21-1606980848-9984276313-839522115-421992
        GroupCount         : 15
        Groups             : {S-1-5-21-1606980848-115178312-839522115-513, S-1-1-0, S-1-5-32-545, S-1-5-4...}
        GroupNames         : -
        Messages           : 

        END  : Look up current user

.EXAMPLE
    PS C:\> Get-PKCurrentUser -LookupGroupNames

        BEGIN: Look up current user and resolve group SIDS to names

        ComputerName       : LAPTOP
        AuthenticationType : Kerberos
        ImpersonationLevel : None
        Name               : DOMAIN\jbloggs
        IsAuthenticated    : True
        IsSystem           : False
        SID                : S-1-5-21-1606980848-9984276313-839522115-421992
        GroupCount         : 15
        Groups             : {S-1-5-21-1606980848-115176313-839522115-513, S-1-1-0, S-1-5-32-545, S-1-5-4...}
        GroupNames         : {Authentication authority asserted identity, BUILTIN\Users, CONSOLE LOGON, 
                             DOMAIN\$222N30-464O03JJ6ILV...}
        Messages           : 

        END  : Look up current user and resolve group SIDS to names

.EXAMPLE
    PS C:\> Get-PKCurrentUser -LookupGroupNames -ExpandGroupNames -Quiet

        ComputerName       : LAPTOP
        AuthenticationType : Kerberos
        ImpersonationLevel : None
        Name               : DOMAIN\jbloggs
        IsAuthenticated    : True
        IsSystem           : False
        SID                : S-1-5-21-1606980848-9984276313-839522115-421992
        GroupCount         : 15
        Groups             : {S-1-5-21-1606980848-115176313-839522115-513, S-1-1-0, S-1-5-32-545, S-1-5-4...}
        GroupNames         : Authentication authority asserted identity
                             BUILTIN\Users
                             CONSOLE LOGON
                             DOMAIN\AtlassianUsers
                             DOMAIN\DBAdmins
                             DOMAIN\chef
                             DOMAIN\ContractEmployeeManagers
                             DOMAIN\Domain Users
                             DOMAIN\Engineering Employee Distro List
                             DOMAIN\EntertainmentCommittee-A
                             DOMAIN\VCenterPowerUsers
                             LOCAL
                             NT AUTHORITY\Authenticated Users
                             NT AUTHORITY\INTERACTIVE
                             NT AUTHORITY\This Organization
        Messages           : 

#>
[CmdletBinding(
    DefaultParameterSetName = "Default"
)]
Param (
    [Parameter(
        ParameterSetName = "Groups",
        HelpMessage = "Attempt lookup of group SIDs to names"
    )]
    [switch]$LookupGroupNames,

    [Parameter(
        ParameterSetName = "Groups",
        HelpMessage = "Expand group names in collection, to stacked strings (presumes -LookupGroupNames)"
    )]
    [switch]$ExpandGroupNames,

    [Parameter(
        HelpMessage = "Suppress non-verbose console output"
    )]
    [switch]$Quiet
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $Source = $PSCmdlet.ParameterSetName
    Switch ($Source) {
        Groups {
            If ($ExpandGroupNames.IsPresent) {$CurrentParams.LookupGroupNames = $LookupGroupNames = $True}
        }
    }
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where{Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"

    #region Load Assembly

    Try {
        $Null = Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    }
    Catch {
        $Msg = "Error loading System.DirectoryServices.AccountManagement assembly"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
        Throw $Msg
    }
 
    #endregion Load Assemblhy

    #region Inner functions

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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    # Function to lookup the group from the SID
    Function GetGroup {
        [CmdletBinding()]
        Param(
            [Parameter(ValueFromPipeline=$True)]
            $Group,
            [Switch]$Expand
        )
        $Output = @()
        Foreach ($G in $Group) {
            Try {
                $Output += ([System.Security.Principal.SecurityIdentifier]$G.value).Translate([system.security.principal.ntaccount])
            }
            Catch {}
        }
        $Output = $Output | Select-Object -Unique | Sort-Object
        If ($Expand.IsPresent) {
            $Output -Join("`n")
        }
        Else {
            $Output
        }
    } #end GetGroup

    #endregion Functions

    # For console
    $Activity = "Look up current user"
    If ($LookupGroupNames.IsPresent) {$Activity += " and resolve group SIDS to names"}
    If ($ExpandGroupNames.IsPresent) {$Activity += " (expand group names to stacked strings)"}
    "BEGIN: $Activity" | Write-MessageInfo -FGColor Yellow -Title

}
Process {
    
    Try {
        $User = [System.Security.Principal.WindowsIdentity]::getcurrent()
        $User | Select @{N="ComputerName";E={$Env:ComputerName}},
            AuthenticationType,
            ImpersonationLevel,
            Name,
            IsAuthenticated,
            IsSystem,
            @{N="SID";E={$_.User}},
            @{N="GroupCount";E={$_.Groups.Count}},
            Groups,
            @{N="GroupNames";E={
                If ($LookupGroupNames.IsPresent) {
                    GetGroup -Group $_.Groups -Expand:$ExpandGroupNames
                }
                Else {"-"}
            }},
            @{N="Messages";E={$Null}}
    }
    Catch {
        Write-Error $_.Exception.Message
        "" | Select @{N="ComputerName";E={$Env:ComputerName}},
        AuthenticationType,
        ImpersonationLevel,
        Name,
        IsAuthenticated,
        IsSystem,
        @{N="SID";E={$Null}},
        @{N="GroupCount";E={$Null}},
        Groups,
        @{N="GroupNames";E={$Null}},
        @{N="Messages";E={$_.Exception.Message}}
    }
        
}
End {
    
    "END  : $Activity" | Write-MessageInfo -FGColor Yellow -Title

}
} #end Get-PKCurrentUser

