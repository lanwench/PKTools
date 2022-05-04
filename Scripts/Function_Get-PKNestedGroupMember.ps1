#requires -version 4
Function Get-PKNestedGroupMembers {
<#
.SYNOPSIS
    Returns direct and nested members for an Active Directory group, including the parent group & depth/level number

.DESCRIPTION
    Returns direct and nested members for an Active Directory group, including the parent group & depth/level number
    Returns basic information (only depth, parent and child distinguishednames, object class, and object name) unless -Detailed is specified 
    Returns a PSObject

.NOTES        
    Name    : Function_Get-PKNestedGroupMembers.ps1
    Created : 2022-04-26
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2022-04-26 - Created script with the invaluable help of Mr. Jeffery Hicks, Super Genius

.PARAMETER Identity
    Active Directory group name or DN

.PARAMETER Depth
    Number of recursion levels for nested group membership expansion (default is all)

.PARAMETER Detailed
    Return additional properties for groups/members

.PARAMETER Server
    Domain controller name or object, or domain FQDN (default is current user's domain)

.PARAMETER Credential
    Valid credentials in domain (default is current user)

.PARAMETER LevelCounter
    FOR INTERNAL USE ONLY - this is an incremental counter used when recursing nested groups

.EXAMPLE
    PS C:\> Get-PKNestedGroupMembers -Identity all_helpdesk -Verbose | Format-Table -AutoSize

        VERBOSE: PSBoundParameters: 
	
        Key              Value                             
        ---              -----                             
        Identity         all_helpdesk
        Verbose          True                              
        Depth            0                                 
        Detailed         False                             
        Server           nlsn.media                        
        Credential                                         
        LevelCounter     0                                 
        PipelineInput    False                             
        ScriptName       Get-PKNestedGroupMembers          
        ScriptVersion    1.0.0                             

        VERBOSE: [Level 0: all_helpdesk] Get Active Directory object for parent group
        VERBOSE: [Level 0: all_helpdesk] Get parent group members
        VERBOSE: [Level 0: all_helpdesk] CN=tier1_helpdesk,OU=Groups,DC=domain,DC=local
        VERBOSE: [Level 0: all_helpdesk] Recursing nested group 'tier1_helpdesk'
        VERBOSE: [Level 1: all_helpdesk] CN=tier2_helpdesk,OU=Groups,DC=domain,DC=local
        VERBOSE: [Level 1: all_helpdesk] Recursing nested group 'tier2_helpdesk'


        Depth ParentDN                                         Name               ObjectClass  DistinguishedName                                                                                 
        ----- --------                                         ----               -----------  -----------------                                                                                 
            0 CN=all_helpdesk,OU=Groups,DC=domain,DC=local     tier1_helpdesk     Group        CN=tier1_helpdesk,OU=Groups,DC=domain,DC=local                 
            1 CN=tier1_helpdesk,OU=Groups,DC=domain,DC=local   Janet Silveira     User         CN=Janet Silveira,OU=Users,DC=domain,DC=local
            1 CN=tier1_helpdesk,OU=Groups,DC=domain,DC=local   Sayeed Nair        User         CN=Sayeed Nair,OU=Users,DC=domain,DC=local
            1 CN=tier1_helpdesk,OU=Groups,DC=domain,DC=local   Vincent Delgado    User         CN=Vincent Delgado,OU=Users,DC=domain,DC=local
            1 CN=all_helpdesk,OU=Groups,DC=domain,DC=local     tier2_helpdesk     Group        CN=tier2_helpdesk,OU=Groups,DC=domain,DC=local                   
            2 CN=tier2_helpdesk,OU=Groups,DC=domain,DC=local   Laura Chu          User         CN=Laura Chu,OU=Users,DC=domain,DC=local
            2 CN=tier2_helpdesk,OU=Groups,DC=domain,DC=local   Kimberley Smith    User         CN=Kimberley Smith,OU=Users,DC=domain,DC=local
            
.EXAMPLE
    PS C:\> Get-PKNestedGroupMembers -Identity datacollection -Detailed

        
        Depth             : 0
        ParentName        : datacollection
        ParentCategory    : Distribution
        ParentScope       : Universal
        ParentMail        : DataCollection@megacorp.com
        ParentDN          : CN=datacollection,OU=Groups,DC=domain,DC=local
        ParentManagedBy   : CN=Janine Curtis,OU=Users,DC=domain,DC=local
        Name              : operations_team
        Mail              : ops@megacorp.com
        ObjectClass       : Group
        UserType          : -
        MemberCount       : 6
        DistinguishedName : CN=operations_team,OU=Groups,DC=domain,DC=local

        Depth             : 1
        ParentName        : operations_team
        ParentCategory    : Distribution
        ParentScope       : Global
        ParentMail        : ops@megacorp.com
        ParentDN          : CN=operations_team,OU=Groups,DC=domain,DC=local
        ParentManagedBy   : CN=Kate Evans,OU=Users,DC=domain,DC=local
        Name              : Pinky Alcazar
        Mail              : pinky@megacorp.com
        ObjectClass       : User
        UserType          : Employee
        MemberCount       : -
        DistinguishedName : CN=Pinky Alcazar,OU=Users,DC=domain,DC=local

        Depth             : 1
        ParentName        : datacollection
        ParentCategory    : Distribution
        ParentScope       : Universal
        ParentMail        : DataCollection@megacorp.com
        ParentDN          : CN=datacollection,OU=Groups,DC=domain,DC=local
        ParentManagedBy   : CN=Janine Curtis,OU=Users,DC=domain,DC=local
        Name              : linux_admins
        Mail              : linuxadmins@megacorp.com
        ObjectClass       : Group
        UserType          : -
        MemberCount       : 38
        DistinguishedName : CN=linux_admins,OU=Groups,DC=domain,DC=local

        Depth             : 2
        ParentName        : linux_admins
        ParentCategory    : Distribution
        ParentScope       : Universal
        ParentMail        : linuxadmins@megacorp.com
        ParentDN          : CN=linux_admins,OU=Groups,DC=domain,DC=local
        ParentManagedBy   : CN=Willa Cather,OU=Users,DC=domain,DC=local
        Name              : Payal Gupta
        Mail              : payal.gupta@megacorp.com
        ObjectClass       : User
        UserType          : Employee
        MemberCount       : -
        DistinguishedName : CN=Payal Gupta,OU=Users,DC=domain,DC=local

        Depth             : 2
        ParentName        : datacollection
        ParentCategory    : Distribution
        ParentScope       : Universal
        ParentMail        : DataCollection@megacorp.com
        ParentDN          : CN=datacollection,OU=Groups,DC=domain,DC=local
        ParentManagedBy   : CN=Janine Curtis,OU=Users,DC=domain,DC=local
        Name              : storage_admins
        Mail              : StorageServices@megacorp.com
        ObjectClass       : Group
        UserType          : -
        MemberCount       : 15
        DistinguishedName : CN=storage_admins,OU=Groups,DC=domain,DC=local

        Depth             : 3
        ParentName        : storage_admins
        ParentCategory    : Distribution
        ParentScope       : Universal
        ParentMail        : StorageServices@megacorp.com
        ParentDN          : CN=storage_admins,OU=ServiceNow,OU=Groups,DC=domain,DC=local
        ParentManagedBy   : CN=Brenda Coliccio,OU=Users,DC=domain,DC=local
        Name              : Aziz Vishnu
        Mail              : Aziz@megacorp.com
        ObjectClass       : User
        UserType          : Employee
        MemberCount       : -
        DistinguishedName : CN=Aziz Vishnu,OU=Users,DC=domain,DC=local

        

#>
[cmdletbinding()]
Param(
    [Parameter(
        Mandatory,
        Position = 0,
        HelpMessage = "Active Directory group name or DN"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Identity,
    
    [Parameter(
        HelpMessage = "Number of recursion levels for nested group membership expansion (default is all)"
    )]
    [ValidateNotNullOrEmpty()]
    [int]$Depth,

    [Parameter(
        HelpMessage = "Return additional properties for groups/members"
    )]
    [Switch]$Detailed,

    [Parameter(
        HelpMessage = "Domain controller name or object, or domain FQDN (default is current user's domain)"
    )]
    [Alias("Domain","ADDomain")]
    [ValidateNotNullOrEmpty()]
    [object]$Server = $env:USERDNSDOMAIN.ToLower(),

    [Parameter(
        HelpMessage = "Valid credentials in domain (default is current user)"
    )]
    [ValidateNotNullOrEmpty()]
    [PSCredential] $Credential,

    [Parameter(
        HelpMessage = "FOR INTERNAL USE ONLY - this is an incremental counter used when recursing nested groups"
    )]
    [ValidateNotNullOrEmpty()]
    [int]$LevelCounter = 0
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $Source = $PSCmdlet.ParameterSetName

    # We need a splat for calling the function within itself

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Splat for Get-AD** 
    $OuterParam = @{
        Server  = $Server
        Verbose = $False
    }
    If ($PSBoundParameters.Credential) {
        $OuterParam.Add("Credential",$Credential)
    }

    # Splat for calling function again within itself, wooo woooooo
    $InnerParam = @{
        Server   = $Server
        Detailed = $CurrentParams.Detailed
        Verbose  = $False
    }
    If ($PSBoundParameters.Credential) {
        $InnerParam.Add("Credential",$Credential)
    }
    If ($CurrentParams.Depth) {
        $InnerParam.Add("Depth",$Depth)
    }

    # Properties to return
    If ($Detailed.IsPresent) {
        $Select = "Depth,ParentName,ParentCategory,ParentScope,ParentMail,ParentDN,ParentManagedBy,Name,Mail,ObjectClass,UserType,MemberCount,DistinguishedName" -split(",")
    }
    Else {
        $Select = "Depth,ParentDN,Name,ObjectClass,DistinguishedName" -split(",")
    }

    # To set title case
    $TextInfo = (Get-Culture).TextInfo

    $Msg = "Get nested group membership" 
    If ($CurrentParams.Depth) {$Msg += ", up to $Depth levels of recursion"}
    Write-Verbose "[BEGIN: $Scriptname] $Msg"
}
Process {
    
    
    Try {
        $Msg = "Get Active Directory object for parent group"
        Write-Verbose "[Level $LevelCounter`: $Identity] $Msg"
        Write-Progress -Activity $Msg -CurrentOperation $Identity -Status "Working"
        
        $GroupObj = Get-ADGroup -Identity $identity -Properties Members,GroupCategory,GroupScope,Mail,ManagedBy @OuterParam -ErrorAction Stop

        Try {
            $Msg = "Get parent group members"
            Write-Verbose "[Level $LevelCounter`: $($GroupObj.Name)] $Msg"

            $Counter = 0
            $Total = $GroupObj.Members.Count

            Foreach ($Member in $GroupObj.Members) {
        
                $Counter ++
                $Msg = $Member
                Write-Verbose "[Level $LevelCounter`: $($GroupObj.Name)] $Msg"
                Write-Progress -Activity "Get nested group membership for group '$($GroupObj.Name)'" -CurrentOperation "$Member" -Status "Level $LevelCounter`: $($GroupObj.Name)" -PercentComplete ($Counter/$Total*100)
            
                $MemberObj = Get-ADObject -Identity $Member -Properties Mail,Extensionattribute11,Member @OuterParam -ErrorAction Stop
                
                $Output = [pscustomobject]@{
                        Depth              = $LevelCounter
                        ParentName         = $GroupObj.Name
                        ParentCategory     = $GroupObj.GroupCategory
                        ParentScope        = $GroupObj.GroupScope
                        ParentMail         = $GroupObj.Mail
                        ParentDN           = $GroupObj.DistinguishedName                    
                        ParentManagedBy    = $GroupObj.ManagedBy
                        Name               = $MemberObj.Name
                        Mail               = $MemberObj.Mail
                        ObjectClass        = $TextInfo.ToTitleCase($MemberObj.ObjectClass)
                        UserType           = &{If ($MemberObj.ObjectClass -eq 'User') {$MemberObj.ExtensionAttribute11} Else {"-"}}
                        MemberCount        = &{If ($MemberObj.ObjectClass -eq "Group") {$MemberObj.Member.Count} Else{"-"}}
                        DistinguishedName  = $MemberObj.DistinguishedName
                    
                    } | Select-Object $Select
                Write-Output $Output

                If ($MemberObj.objectclass -eq 'Group') {
                    If ($CurrentParams.Depth) {
                        While ($LevelCounter -lt $Depth) {
                            $Msg = "Recursing nested group '$($MemberObj.Name)'"
                            Write-Verbose "[Level $LevelCounter`: $($GroupObj.Name)] $Msg"
                            $LevelCounter = $($LevelCounter +1)
                            Get-PKNestedGroupMembers -Identity $Member @InnerParam -Level $LevelCounter
                        }
                        If ($LevelCounter -eq $Depth) {
                            $Msg = "Maximum recursion depth of $Depth has been reached"                            
                            Write-Warning "[Level $LevelCounter`: $($GroupObj.Name)] $Msg"
                        }
                    }
                    Else {
                        $Msg = "Recursing nested group '$($MemberObj.Name)'"
                        Write-Verbose "[Level $LevelCounter`: $($GroupObj.Name)] $Msg"
                        $LevelCounter = $($LevelCounter +1)
                        Get-PKNestedGroupMembers -Identity $Member @InnerParam -Level $LevelCounter
                    }
                }
            } #end foreach
        }
        Catch {
            $Msg = "An error occurred $($_.Exception.Message)"
            Write-Warning "[Level $LevelCounter`: $($GroupObj.Name)] $Msg"
        }
    }
    Catch {
        $Msg = "An error occurred $($_.Exception.Message)"
        Write-Warning "[Level $LevelCounter`: $Identity] $Msg"
    }
}
End {
    $Null = Write-Progress -Activity * -Completed
}
} #end Get-PKNestedGroupMembers
