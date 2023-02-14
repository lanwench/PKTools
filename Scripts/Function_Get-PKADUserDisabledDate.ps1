#requires -version 4
Function Get-PKADUserDisabledDate {
<#
.SYNOPSIS 
    Uses Get-ADUser and Get-ADReplicationAttributeMetadata to return the date user objects were disabled

.DESCRIPTION
    Uses Get-ADUser and Get-ADReplicationAttributeMetadata to return the date user objects were disabled
    Accepts identity or searchbase (with choice of searchscope)
    Returns a PSObject

.NOTES
    Name    : Get-PKADUserDisabledDate 
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2023-02-14 - Created script

.EXAMPLE
    PS C:\> Get-PKADUserDisabledDate -Identity jbloggs -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                   
        ---           -----                   
        Identity      {jbloggs}           
        Verbose       True                    
        Searchbase                            
        SearchScope   Subtree                 
        Detailed      False                   
        Server                                
        Credential                            
        ScriptName    Get-PKADUserDisabledDate
        ScriptVersion 1.0.0                   
        PipelineInput False                   

        VERBOSE: [Prerequisites] Getting domain controller
        VERBOSE: [jbloggs] Getting user object
        VERBOSE: [jbloggs] Getting disabled date

        FullName          : Joe Bloggs
        Enabled           : False
        WhenDisabled      : 2023-01-11 14:01:46Z
        SAMAccountName    : jbloggs
        Mail              : joe.bloggs@megacorp.com
        DistinguishedName : CN=Joe Bloggs,OU=Disabled,OU=Accounts,DC=megacorp,DC=local

.EXAMPLE
    PS C:\> Get-ADUser -SearchBase "OU=Accounts,DC=megacorp,DC=local" -Filter "Enabled -eq '$False'" | Get-PKADUserDisabledDate -Server dc.megacorp.local -Credential $MyCred | Format-Table -AutoSize

        FullName             Enabled WhenDisabled        SAMAccountName  Mail DistinguishedName                                                                                 
        --------             ------- ------------        --------------  ---- -----------------                                                                                 
        Marybeth Sanderson   False 2023-01-30 17:41:45Z  msanderson      CN=Marybeth Sanderson,OU=Disabled,OU=Accounts,DC=megacorp,DC=local
        Bob Coughlan         False 2022-06-13 21:25:49Z  bcoughlan       CN=Bob Coughlan,OU=Contractors,OU=Accounts,DC=megacorp,DC=local
        Brad Taylor          False 2023-01-03 08:02:17Z  btaylor         CN=Brad Taylor,OU=Enabled,OU=Accounts,DC=megacorp,DC=local
        
.EXAMPLE
    PS C:\> Get-PKADUserDisabledDate -Identity msanderson,kbates -Detailed

        FullName          : Marybeth Sanderson
        Name              : Marybeth Sanderson
        Enabled           : False
        WhenDisabled      : 2023-01-30 17:41:45Z
        SAMAccountName    : msanderson
        WhenCreated       : 2020-07-07 09:10:44Z
        WhenChanged       : 2023-01-30 17:53:56Z
        Manager           : CN=Brad Taylor,OU=Enabled,OU=Accounts,DC=megacorp,DC=local
        CanonicalName     : megacorp.local/Accounts/Enabled/Marybeth Sanderson
        DistinguishedName : CN=Marybeth Sanderson,OU=Enabled,OU=Accounts,DC=megacorp,DC=local

        WARNING: [kbates] User object is not currently disabled
        

#>
[CmdletBinding(DefaultParameterSetName="Identity")]
Param(
    [Parameter(
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        ParameterSetName = "Identity",
        Mandatory
    )]
    [object[]]$Identity,

    [Parameter(
        ParameterSetName = "Searchbase",
        Mandatory
    )]
    [object]$Searchbase,

    [Parameter(
        ParameterSetName = "Searchbase"
    )]
    [ValidateSet("Subtree","Base","OneLevel")]
    [object]$SearchScope = "Subtree",

    [Parameter()]
    [switch]$Detailed,

    [Parameter()]
    [object]$Server,

    [Parameter()]
    [pscredential]$Credential
)
Begin{
    
     
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here?
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $Source = $PSCmdlet.ParameterSetName

    $CurrentParams = $PSBoundParameters
    $ScriptName = $MyInvocation.MyCommand.Name
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path Variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Params = @{
        ErrorAction = "Stop"
        Verbose = $False
    }
    If ($CurrentParams.Credential) {
        $Params.Add("Credential",$Credential)
    }

    If (-not $CurrentParams.Server) {
        $Msg = "Getting domain controller"
        Write-Verbose "[Prerequisites] $Msg"
        Write-Progress -Activity "Prerequisites" -CurrentOperation $Msg 
        Try {
            $DC = (Get-ADDomain (Get-ADForest @Params).RootDomain @Params).PDCEmulator
            $Params.Add("Server",$DC)
        }
        Catch {
            Throw $_.Exception.Message
        }
    }
    
    Else {
        $DC = $Server
        $Params.Add("Server",$DC)
        <#
        $Msg = "Verifying domain controller"
        Write-Verbose "[Prerequisites] $Msg"
        Write-Progress -Activity "Prerequisites" -CurrentOperation $Msg 
        Try {
            $DC = (Get-ADDomainController -Identity $Server -Server $Server @Params).Hostname
            $Params.Add("Server",$DC)
        }
        Catch {
            Throw $_.Exception.Message
        }
        #>
    }

    If ($Detailed.IsPresent) {
        $Select = "FullName,Name,Enabled,WhenDisabled,SAMAccountName,WhenCreated,WhenChanged,Mail,Manager,CanonicalName,DistinguishedName" -split(",")
    }
    Else {
        $Select = "FullName,Enabled,WhenDisabled,SAMAccountName,Mail,DistinguishedName" -split(",")
    }


    $Props = "GivenName,Surname,WhenChanged,Mail,Name,Manager,Description,WhenCreated,CanonicalName" -split(",")

    $Activity = "Report date user objects were disabled in Active Directory"
}
Process {
    
    Switch ($Source) {
        Identity {
            $Total = $Identity.Count
            $Current = 0
            Foreach ($User in $Identity) {
                $Current ++
                
                Try {
                    
                    If ($User -is [Microsoft.ActiveDirectory.Management.ADAccount]) {$UserObj = $User}
                    Else {
                        $Msg = "Getting user object"
                        Write-Verbose "[$User] $Msg"
                        Write-Progress -Activity $Activity -CurrentOperation $User -Status $Msg -PercentComplete ($Current/$Total*100)
                        $UserObj = Get-Aduser -Identity $User -Properties $Props
                    }


                    If (-not $UserObj.Enabled) {
                        Try {
                            $Msg = "Getting disabled date"
                            Write-Verbose "[$User] $Msg"
                            $UserObj | Get-ADReplicationAttributeMetadata -Properties UserAccountControl @Params | 
                                Select-Object @{N="FullName";E={"$($UserObj.GivenName) $($UserObj.Surname)"}},
                                    @{N="Name";E={$UserObj.Name}},
                                    @{N="Enabled";E={$UserObj.Enabled}},
                                    @{N="WhenDisabled";E={Get-Date ($_.LastOriginatingChangeTime) -Format u}},
                                    @{N="SAMAccountName";E={$UserObj.SAMAccountName}},
                                    @{N="WhenCreated";E={Get-Date ($UserObj.WhenCreated) -Format u}},
                                    @{N="WhenChanged";E={Get-Date ($UserObj.WhenChanged) -Format u}},
                                    @{N="Mail";E={$UserObj.Mail}},
                                    @{N="Description";E={$UserObj.Description}},
                                    @{N="Manager";E={$UserObj.Manager}},
                                    @{N="CanonicalName";E={$UserObj.CanonicalName}},
                                    @{N="DistinguishedName";E={$UserObj.DistinguishedName}} | Select-Object $Select
                        }

                        

                        Catch {
                            Throw "[$User] $($_.Exception.Message)"
                        }
                    }
                    Else {
                        $Msg = "User object is not currently disabled"
                        Write-Warning "[$User] $Msg"
                    }
                }
                Catch {
                    Throw "[$User] $($_.Exception.Message)"
                }
            
            }
        }
        Searchbase {
            Try {
                $Msg = "Getting disabled users in searchbase"
                Write-Verbose "[$Searchbase] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Searchbase -Status $Msg
                Get-Aduser -Searchbase $Searchbase -SearchScope $SearchScope -filter "Enabled -eq '$False'" -Properties $Props @Params -PipelineVariable UserObj | 
                    Get-ADReplicationAttributeMetadata -Properties UserAccountControl @Params | 
                        Select-Object @{N="FullName";E={"$($UserObj.GivenName) $($UserObj.Surname)"}},
                            @{N="Name";E={$UserObj.Name}},
                            @{N="Enabled";E={$UserObj.Enabled}},
                            @{N="WhenDisabled";E={Get-Date ($_.LastOriginatingChangeTime) -Format u}},
                            @{N="SAMAccountName";E={$UserObj.SAMAccountName}},
                            @{N="WhenCreated";E={Get-Date ($UserObj.WhenCreated) -Format u}},
                            @{N="WhenChanged";E={Get-Date ($UserObj.WhenChanged) -Format u}},
                            @{N="Mail";E={$UserObj.Mail}},
                            @{N="Description";E={$UserObj.Description}},
                            @{N="Manager";E={$UserObj.Manager}},
                            @{N="CanonicalName";E={$UserObj.CanonicalName}},
                            @{N="DistinguishedName";E={$UserObj.DistinguishedName}} | Select-Object $Select
            
            }
            Catch {
                Throw "[$Searchbase] $($_.Exception.Message)"
            }
        }

    
    }
    <#

    $Members = Get-Aduser -Searchbase "DC=nlsn,DC=media" -SearchScope Subtree -filter "Enabled -eq '$False'"  -Server daynlsndc-1.nlsn.media -Properties WhenChanged,Mail,Name,Manager,Description,WhenCreated,CanonicalName -PipelineVariable User | 
    Get-ADReplicationAttributeMetadata -Server daynlsndc-1.nlsn.media -Properties UserAccountControl  | 
        Select-Object @{N="Name";E={$User.Name}},
        @{N="Enabled";E={$User.Enabled}},
        @{N="SAMAccountName";E={$User.SAMAccountName}},
        @{N="Mail";E={$User.Mail}},
        @{N="Description";E={$User.Description}},
        @{N="Manager";E={$User.Manager}},
        @{N="WhenChanged";E={$User.WhenChanged}},
        @{N="WhenDisabled";E={$User.LastOriginatingChangeTime}},
        @{N="CanonicalName";E={$User.CanonicalName}},
        @{N="DistinguishedName";E={$User.DistinguishedName}}


        #>
}
}

