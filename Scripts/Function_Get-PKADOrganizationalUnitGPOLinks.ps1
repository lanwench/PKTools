#requires -Version 3
Function Get-PKADOrganizationalUnitLinkedGPOs {
<#
.SYNOPSIS
    Gets details or counts of Group Policy object linked to Active Directory organizational units

.DESCRIPTION
    Gets details or counts of Group Policy object linked to Active Directory organizational units
    Optionally returns number of objects within OU
    Returns a PSObject

.NOTES
    Name    : Function_Get-PKADOrganizationalUnitLinkedGPOs.ps1
    Author  : Paula Kingsley
    Created : 2019-02-22
    Version : 01.00.0000
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2019-02-22 - Created script


.EXAMPLE
    PS C:\> Get-PKADOrganizationalUnitLinkedGPOs -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                               
        ---                   -----                               
        ADDomain              domain.local             
        Verbose               True                                
        OU                                                        
        GetOUMemberCount      False                               
        Server                                                    
        UseCanonicalName      False                               
        CNSplitLevel          0                                   
        Credential                                                
        SuppressConsoleOutput False                               
        SearchFields                                              
        ParameterSetName      OU                                  
        PipelineInput         True                                
        ScriptName            Get-PKADOrganizationalUnitLinkedGPOs
        ScriptVersion         1.0.0                               


        VERBOSE: [Prerequisites] Connect to Active Directory
        VERBOSE: [Prerequisites] Successfully connected to Active Directory Domain 'domain.local'
        VERBOSE: [Prerequisites] Find nearest Active Directory Domain Controller
        VERBOSE: [Prerequisites] Successfully connected to Active Directory Domain Controller 'dc02.domain.local'
        SCRIPT : Get linked Group Policy object names, counts, and computer objects in organizational unit
        VERBOSE: Get Organizational Unit(s)
        VERBOSE: 465 Organizational Unit object(s) found
        VERBOSE: domain.local/Corp
        VERBOSE: domain.local/Corp/All Groups
        VERBOSE: domain.local/Corp/All Groups/Security Groups
        VERBOSE: domain.local/Microsoft Exchange Security Groups
        VERBOSE: domain.local/Corp/Administrative
        VERBOSE: domain.local/Corp/Administrative/Service Accounts
        VERBOSE: domain.local/Domain Controllers
        VERBOSE: domain.local/Corp/All Users
        VERBOSE: domain.local/Corp/Computers
        VERBOSE: domain.local/Corp/Computers/Servers/
        VERBOSE: domain.local/Corp/Computers/Servers/Infra
        VERBOSE: domain.local/Corp/Computers/Servers/Production
        VERBOSE: domain.local/Corp/Computers/Workstations
        VERBOSE: domain.local/Corp/Computers/Workstations/Japan
        VERBOSE: domain.local/Corp/Computers/Workstations/Munich

        <snip>

        OUName            : Telecom
        MemberObjCount    : -
        GPOCount          : 2
        GPONames          : {ops_prod_telecom_svcaccounts, ops_prod_telecom_admins}
        CanonicalName     : domain.local/Corp/Computers/Servers/Telecom
        DistinguishedName : OU=Telecom,OU=Servers,OU=Computers,OU=Corp,DC=domain,DC=local

        OUName            : SQL
        MemberObjCount    : -
        GPOCount          : 0
        GPONames          : -
        CanonicalName     : domain.local/Corp/Computers/Servers/Telecom/SQL
        DistinguishedName : OU=SQL,OU=Telecom,OU=Servers,OU=Computers,OU=Corp,DC=domain,DC=local

        OUName            : WebProd
        MemberObjCount    : -
        GPOCount          : 2
        GPONames          : {ops_dev_web_svcaccounts, ops_dev_web_admins, disable_remoteps}
        CanonicalName     : domain.local/Corp/Computers/Servers/WebProd
        DistinguishedName : OU=WebProd,OU=Servers,OU=Computers,OU=Corp,DC=domain,DC=local

        OUName            : Lab
        MemberObjCount    : -
        GPOCount          : 6
        GPONames          : {Set workstation Local Administrator Password, Microsoft Office 2010 x64, Disable SMBv1 }
        CanonicalName     : domain.local/Corp/Computers/Servers/Lab
        DistinguishedName : OU=Lab,OU=Servers,OU=Computers,OU=Corp,DC=domain,DC=local

        <snip>

.EXAMPLE
    PS C:\> Get-PKADOrganizationalUnitLinkedGPOs -OU "OU=Test,OU=Toledo,OU=Workstations,OU=Computers,OU=Sandbox,OU=Dev,DC=foo,DC=com" -ADDomain foo.com -Credential $Cred -Server mydc.foo.com -GetOUMemberCount -UseCanonicalName -CNSplitLevel 4
        
        ACTION : Get linked Group Policy object names and counts, and member object counts, in organizational unit

        OUName            : Workstations/Toledo/Test
        MemberObjCount    : 14
        GPOCount          : 2
        GPONames          : {Disable SMBv1 , it_general_winrm_enabled}
        CanonicalName     : foo.com.com/Dev/Sandbox/Computers/Workstations/Toledo/Test
        DistinguishedName : OU=Test,OU=Toledo,OU=Workstations,OU=Computers,OU=Sandbox,OU=Dev,DC=foo,DC=com
#>
[CmdletBinding(DefaultParameterSetName = "Default")]
Param(
    
    [Parameter(
        HelpMessage = "NetBIOS name or FQDN of Active Directory Domain (default is current user's)"
    )]
    [ValidateNotNullOrEmpty()]
    [string] $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,

    [Parameter(
        HelpMessage = "Organizational unit distinguishedname (default is all OUs at root of domain)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$OU,

    [Parameter(
        HelpMessage="Get object count for OU members"
    )]
    [Switch]$GetOUMemberCount,

    [Parameter(
        HelpMessage = "NetBIOS name or FQDN name of Domain Controller (default is first available)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [Parameter(
        ParameterSetName = "CN",
        HelpMessage = "Use CN attribute for Organizational Unit name in output, not Name"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$UseCanonicalName,

    [Parameter(
        ParameterSetName = "CN",
        HelpMessage = "Split CanonicalName as OU name, up to n previous levels"
    )]
    [ValidateNotNullOrEmpty()]
    [int]$CNSplitLevel = 0,

    [Parameter(
        HelpMessage="AD credentials (default is current user)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential,

    [Parameter(
        HelpMessage="Suppress non-verbose console output"
    )]
    [Switch]$SuppressConsoleOutput


)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("Name")) -and (-not $Name)
    
    # Display our parameters
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.SearchFields = $CurrentParams.SearchFields | Where-Object {$_ -notmatch "AllListed"}
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #endregion Show parameters

    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"

    #region Test for AD module

    If (-not ($Null = Get-Module ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue)) {
        $Msg = "This function requires the ActiveDirectory module"
        $Host.UI.WriteErrorLine("ERROR  : $Msg")
        Break
    }

    #endregion Test for AD module


    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    # Splat for Get-ADDomain
    $Param_GetAD = @{}
    $Param_GetAD = @{
        Identity    = $ADDomain
        ErrorAction = "Stop"
        Verbose     = $False
    }
    If ($CurrentParams.Credential) {
        $Param_GetAD.Add("Credential",$Credential)
    }

    # Splat for Get-ADDomainController
    If ($CurrentParams.Server) {
        $Param_GetDC = @{}
        $Param_GetDC = @{
            Identity    = $Server
            Server      = $Server
            ErrorAction = "Stop"
            Verbose     = $False
        }
        If ($CurrentParams.Credential) {
            $Param_GetDC.Add("Credential",$Credential)
        }
    }
    Else {
        $Param_GetDC = @{}
        $Param_GetDC = @{
            DomainName      = $Null
            Discover        = $True
            NextClosestSite = $True
            ErrorAction     = "Stop"
            Verbose         = $False
        }
    }

    # Splat for Get-ADOrganizationalUnit
    $Param_GetOU = @{}
    $Param_GetOU = @{
        Server      = $Null
        Filter      = "*"
        SearchBase  = $Null
        SearchScope = "Subtree"
        Properties  = "Name","CanonicalName","DistinguishedName","gpLink"
        ErrorAction = "SilentlyContinue"
        Verbose     = $False
    }
    If ($CurrentParams.Credential){
        $Param_GetOU.Add("Credential",$Credential)
    }
    
    #Splat for Get-ADComputer
    $Param_GetOUMember = @{}
    $Param_GetOUMember = @{
        Searchbase  = $Null
        SearchScope = "OneLevel"
        Filter      = "*"
        Server      = $Null
        ErrorAction = "SilentlyContinue"
        Verbose     = $False
    }
    If ($CurrentParams.Credential){
        $Param_GetOUMember.Add("Credential",$Credential)
    }

    # Splat for Write-Progress
    $Activity1 = "Get AD organizational unit(s)"
    $Param_WP1 = @{}
    $Param_WP1 = @{
        Activity         = $Activity1
        CurrentOperation = $Null
        Status           = "Working"
    }

    # Splat for Write-Progress (inner)
    If ($GetOUMemberCount.IsPresent) {$Activity2 = "Get linked Group Policy object names and counts, and member object counts"}
    Else {$Activity2 = "Get linked Group Policy object names and counts"}
    $Param_WP2 = @{}
    $Param_WP2 = @{
        Activity         = $Activity2
        CurrentOperation = $Null
        ID               = 1
        Status           = "Working"
        PercentComplete  = $Null
    }

    #endregion Splats

    #region Prerequisites 

    # Connect to AD
    $Activity = "Prerequisites"
    $Msg = "Connect to Active Directory"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    Try {
        
        $ADConfirm = Get-ADDomain @Param_GetAD
        $Msg = "Successfully connected to Active Directory Domain '$($ADConfirm.DNSRoot.Tolower())'"
        Write-Verbose "[Prerequisites] $Msg"
        
        # Get the domain controller
        If (-not $CurrentParams.Server) {
            
            $Msg = "Find nearest Active Directory Domain Controllerr"
            Write-Verbose "[Prerequisites] $Msg"
            
            Try {
                $Param_GetDC.DomainName = $ADConfirm.DNSRoot
                $DCObj = Get-ADDomainController @Param_GetDC
                $DC = $($DCObj.HostName)
                $Msg = "Successfully connected to Active Directory Domain Controller '$DC'"
                Write-Verbose "[Prerequisites] $Msg"
            }
            Catch {
                $Msg = "Failed to find Domain Controller for Active Directory Domain '$($ADConfirm.DNSRoot.Tolower())'"
                If ($ErrorDetails = $_.exception.message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                Break
            }    
        }
        Else {
            $Msg = "Connect to to Active Directory Domain Controller"
            Write-Verbose "[Prerequisites] $Msg"

            Try {
                $DCObj = Get-ADDomainController @Param_GetDC
                If ($DCObj.Domain -eq $ADConfirm.DNSRoot) {
                    $DC = $($DCObj.HostName)
                    $Msg = "Successfully connected to Active Directory Domain Controller '$DC'"
                    Write-Verbose "[Prerequisites] $Msg"
                }
                Else {
                    $Msg = "Domain Controller '$($DCObj.HostName)' is not in Active Directory Domain '$($ADConfirm.DNSRoot)'"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    Break
                }
            }
            Catch {
                $Msg = "Failed to find Domain Controller'$Server' in Active Directory Domain '$($ADConfirm.DNSRoot.Tolower())'"
                If ($ErrorDetails = $_.exception.message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                Break
            }
        }    
        
        # Update the splat
        If ($DCObj) {
            $Param_GetOU.Server = $Param_GetOUMember.Server = $DC
        }

    }
    Catch [exception] {
        $Msg = "Failed to connect to Active Directory Domain '$ADDomain'"
        If ($ErrorDetails = $_.exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
        Break
    }

    Write-Progress -Activity $Activity -Completed

    #endregion Prerequisites

    # Console output
    If ($GetOUMemberCount.IsPresent) {$Activity = "Get linked Group Policy object names and counts, and member object counts, in organizational unit"}
    Else {$Activity = "Get linked Group Policy object names and counts in organizational unit"}
    $BGColor = $Host.UI.RawUI.BackgroundColor
    $Msg = "ACTION : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}

Process {
    
    # Get the OU object(s)
    If (-not $CurrentParams.OU) {
        $Param_GetOU.Searchbase = $ADConfirm.DistinguishedName
    }
    Else {
        $Param_GetOU.SearchBase = $OU
    }
    
    Try {
        $Msg = "Get Organizational Unit(s)"
        Write-Verbose $Msg
        $Param_WP1.CurrentOperation = $Msg
        Write-Progress @Param_WP1
        $OUObj = Get-ADOrganizationalUnit @Param_GetOU
     
        If (($OUObj -as [array]).Count -gt 0) {

            Try {
        
                $Total = ($OUObj -as [array]).Count
                $Current = 0

                $Msg = "$Total Organizational Unit object(s) found"
                Write-Verbose $Msg

                Foreach ($Obj in $OUObj) {
            
                    $Current ++
                    Write-Verbose $Obj.CanonicalName
                    $Param_WP2.CurrentOperation = $Obj.CanonicalName
                    $Param_WP2.PercentComplete = ($Current/$Total*100)    
                    Write-Progress @Param_WP2
                
                    # Format list of GPOs
                    [array]$GPO = $Obj.LinkedGroupPolicyObjects | ForEach-Object {([adsi]"LDAP://$_").displayName -join ''}
                    
                    # Get count of computers in OU
                    $Param_GetOUMember.SearchBase = $Obj.DistinguishedName
                    [array]$Members = Get-ADObject @Param_GetOUMember

                    Switch ($UseCanonicalName) {
                        False {
                            $OUName = $Obj.Name
                        }
                        True  {
                            If ($CNSplitLevel -gt 0) {
                                $OUName = $Obj.CanonicalName -Split("/")
                                If ($OUName.Count -gt $CNSplitLevel) {
                                    $OUName = $OUName[$CNSplitLevel..$OUName.Count] -join("/")
                                }
                                Else {$OUName = $Obj.CanonicalName}
                            }
                        }
                    }
                    
                    # Output object
                    New-Object PSObject -Property([ordered] @{
                        OUName            = $OUName
                        MemberObjCount    = &{
                            If ($GetOUMemberCount.IsPresent) { $Computers.Count}
                            Else {"-"}
                                            }
                        GPOCount          = $GPO.Count
                        GPONames          = &{
                            #If ($GPO.Count -gt 0) {$GPO -join("; ")}
                            If ($GPO.Count -gt 0) {$GPO | Where-Object {$_}}
                            Else {"-"}
                        }
                        CanonicalName     = $Obj.CanonicalName
                        DistinguishedName = $Obj.DistinguishedName
                    }) 
                } #end foreach object

            }
            Catch {
                $Msg = "Failed to get linked GPOs and computer object counts under searchbase '$($Obj.DistinguishedName)'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "" ; $ErrorDetails}
                $Host.UI.WriteErrorLine("ERROR  : $Msg")
            }
        }

        Else {
            $Msg = "Failed to get OrganizationalUnit(s) under searchbase '$($Param_GetOU.SearchBase)'"
            $Host.UI.WriteErrorLine("ERROR  : $Msg")
        }
    }
    Catch {
        $Msg = "Failed to get OrganizationalUnit(s) under searchbase '$($Param_GetOU.SearchBase)'"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "" ; $ErrorDetails}
        $Host.UI.WriteErrorLine("ERROR  : $Msg")
    }
    

}
End {
    $Null = Write-Progress -Activity * -Completed
}
} #end function