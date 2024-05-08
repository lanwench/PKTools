
#requires -Version 4

Function Get-PKADDisabledObject {
    <#
.SYNOPSIS
    Retrieves details of disabled Active Directory objects

.DESCRIPTION
    The Get-PKADDisabledObject function retrieves details of disabled Active Directory objects, such as users, computers, service accounts, or any object class. 
    It takes one or more Active Directory objects or object names as input and returns information about the objects, including the date the object was last disabled.
    Because this supports string input, it uses Where-Object to filter out enabled objects. 
    This is less efficient than using the -Filter parameter, but we can't specify names AND filter, so we have to use the slower method. 
    To speed things up, ensure the input consists only of disabled objects.

.PARAMETER Name
    Specifies one or more object names. This parameter supports pipeline input.

.PARAMETER ObjectClass
    Specifies the object class to filter the search. Valid values are "Any" (default), "User", "Computer", or "ServiceAccount".

.PARAMETER Server
    Specifies the name of the domain controller (PDC emulator) to use for the search.

.PARAMETER Credential
    Specifies alternate credentials to use for the search.

.INPUTS
    System.Object

.OUTPUTS
    System.Management.Automation.PSCustomObject

.EXAMPLE
    PS C:\> Get-PKADDisabledObject -Name "JohnDoe" -ObjectClass "User"
    Retrieves details of the disabled Active Directory user with the name "JohnDoe".

.EXAMPLE
    PS C:\> "JohnDoe", "JaneSmith" | Get-PKADDisabledObject -ObjectClass "User"
    Retrieves details of the disabled Active Directory users with the names "JohnDoe" and "JaneSmith".

.NOTES
    Name    : Function_Get-PKADDisabledObject.ps1
    Created : 2024-05-08
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2024-05-08 - Created script
#>
    [CmdletBinding()]
    Param(
        [Parameter(
            Position = 0,
            HelpMessage = "One or more object names",
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Mandatory
        )][Object[]]$Name,

        [Parameter(
            HelpMessage = "Object class: user, computer, service account, or any (default is Any)"
        )]
        [ValidateSet("Any", "User", "Computer", "ServiceAccount")]
        [ValidateNotNullOrEmpty()]
        [string]$ObjectClass = "Any",

        [Parameter(
            HelpMessage = "Name of domain controller (PDC emulator)"
        )]
        [ValidateNotNullOrEmpty()]
        [Object[]]$Server,

        [Parameter(
            HelpMessage = "Alternate credentials"
        )]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential
    )
    Begin {
        
        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # Show our settings
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
            Where-Object { Test-Path variable:$_ } | Foreach-Object {
                $CurrentParams.Add($_, (Get-Variable $_).value)
            }
        $CurrentParams.Add("PipelineInput", $PipelineInput)
        $CurrentParams.Add("ScriptName", $MyInvocation.MyCommand.Name)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
        $Param = @{Verbose = $False }
        If ($PSBoundParameters.ContainsKey($Credential)) {
            $Param['Credential'] = $Credential
        }
        If (-not ($PSBoundParameters.ContainsKey($Server))) {
            $Param['Server'] = ((Get-ADDomain @Param (Get-ADForest @Param).RootDomain).PDCEmulator)
        }

        Switch ($ObjectClass) {
            Any {
                $Properties = "Name,SAMAccountName,Enabled,MemberOf,WhenChanged,WhenCreated,UserAccountControl,LastLogonTimeStamp,Manager,CanonicalName,DistinguishedName" -split (",")
                $Select = @{N = "Name"; E = { $Obj.Name } },
                @{N = "Enabled"; E = { ($Obj.userAccountControl -band 0x0002) -eq 0 } },
                @{N = "ObjectClass"; E = { $Obj.ObjectClass } },
                @{N = "SAMAccountName"; E = { $Obj.SAMAccountName } },
                @{N = "WhenDisabled"; E = { Get-Date ($UAC.LastOriginatingChangeTime) -Format u } },
                @{N = "WhenCreated"; E = { Get-Date ($Obj.WhenCreated) -Format u } },
                @{N = "WhenChanged"; E = { Get-Date ($Obj.WhenChanged) -Format u } },
                @{N = "LastLogonTimestamp"; E = { Get-Date $([datetime]::FromFileTime($Obj.LastLogonTimestamp)) -Format u } },
                @{N = "Description"; E = { $Obj.Description } },
                @{N = "Manager"; E = { $Obj.Manager -replace "(CN=)(.*?),.*", '$2' } },
                @{N = "CanonicalName"; E = { $Obj.CanonicalName } },
                @{N = "DistinguishedName"; E = { $Obj.DistinguishedName } }
            }
            User {
                $Properties = "Name,SAMAccountName,Enabled,MemberOf,WhenChanged,WhenCreated,LastLogonDate,Manager,UserAccountControl,CanonicalName,DistinguishedName" -split (",")
                $Select = @{N = "Name"; E = { $Obj.Name } },
                @{N = "Enabled"; E = { $Obj.Enabled } },
                @{N = "ObjectClass"; E = { $Obj.ObjectClass } },
                @{N = "SAMAccountName"; E = { $Obj.SAMAccountName } },
                @{N = "WhenDisabled"; E = { Get-Date ($UAC.LastOriginatingChangeTime) -Format u } },
                @{N = "WhenCreated"; E = { Get-Date ($Obj.WhenCreated) -Format u } },
                @{N = "WhenChanged"; E = { Get-Date ($Obj.WhenChanged) -Format u } },
                @{N = "LastLogonDate"; E = { Get-Date ($Obj.LastLogonDate) -Format u } },
                @{N = "NumGroups"; E = { $Obj.MemberOf.Count } },
                @{N = "Description"; E = { $Obj.Description } },
                @{N = "Manager"; E = { $Obj.Manager -replace "(CN=)(.*?),.*", '$2' } },
                @{N = "CanonicalName"; E = { $Obj.CanonicalName } },
                @{N = "DistinguishedName"; E = { $Obj.DistinguishedName } }
            }
            Computer {
                $Properties = "Name,SAMAccountName,Enabled,MemberOf,WhenChanged,WhenCreated,LastLogonDate,ManagedBy,OperatingSystem,UserAccountControl,IPv4Address,CanonicalName,DistinguishedName" -split (",")
                $Select = @{N = "Name"; E = { $Obj.Name } },
                @{N = "Enabled"; E = { $Obj.Enabled } },
                @{N = "ObjectClass"; E = { $Obj.ObjectClass } },
                @{N = "SAMAccountName"; E = { $Obj.SAMAccountName } },
                @{N = "WhenDisabled"; E = { Get-Date ($UAC.LastOriginatingChangeTime) -Format u } },
                @{N = "WhenCreated"; E = { Get-Date ($Obj.WhenCreated) -Format u } },
                @{N = "WhenChanged"; E = { Get-Date ($Obj.WhenChanged) -Format u } },
                @{N = "LastLogonDate"; E = { Get-Date ($Obj.LastLogonDate) -Format u } },
                @{N = "OperatingSystem"; E = { $Obj.OperatingSystem } },
                @{N = "IPv4Address"; E = { $Obj.IPv4Address } },
                @{N = "Description"; E = { $Obj.Description } },
                @{N = "ManagedBy"; E = { $Obj.ManagedBy -replace "(CN=)(.*?),.*", '$2' } },
                @{N = "CanonicalName"; E = { $Obj.CanonicalName } },
                @{N = "DistinguishedName"; E = { $Obj.DistinguishedName } }
            }
            ServiceAccount {
                $Properties = "Name,SAMAccountName,Enabled,MemberOf,WhenChanged,WhenCreated,LastLogonDate,Manager,UserAccountControl,CanonicalName,DistinguishedName" -split (",")
                $Select = @{N = "Name"; E = { $Obj.Name } },
                @{N = "Enabled"; E = { $Obj.Enabled } },
                @{N = "ObjectClass"; E = { $Obj.ObjectClass } },
                @{N = "SAMAccountName"; E = { $Obj.SAMAccountName } },
                @{N = "WhenDisabled"; E = { Get-Date ($UAC.LastOriginatingChangeTime) -Format u } },
                @{N = "WhenCreated"; E = { Get-Date ($Obj.WhenCreated) -Format u } },
                @{N = "WhenChanged"; E = { Get-Date ($Obj.WhenChanged) -Format u } },
                @{N = "LastLogonDate"; E = { Get-Date ($Obj.LastLogonDate) -Format u } },
                @{N = "NumGroups"; E = { $Obj.MemberOf.Count } },
                @{N = "Description"; E = { $Obj.Description } },
                @{N = "Manager"; E = { $Obj.Manager -replace "(CN=)(.*?),.*", '$2' } },
                @{N = "CanonicalName"; E = { $Obj.CanonicalName } },
                @{N = "DistinguishedName"; E = { $Obj.DistinguishedName } }
            }
        }

        $Activity = "Report details of disabled Active Directory objects"
        Write-Verbose "[BEGIN: $ScriptName] $Activity"

    }
    Process {

        $Total = $Name.Count
        $Current = 0
        Foreach ($i in $Name) {
        
            $Current ++
            Try { 
                Switch ($ObjectClass) {
                    User {
                        $Msg = "Looking for Active Directory user"
                        Write-Verbose "[$i] $Msg"
                        Write-Progress -Activity $Activity -CurrentOperation $i -Status "ObjectClass: $ObjectClass" -PercentComplete ($Current / $Total * 100)
                        $Obj = Get-ADUser $I @Param -Properties $Properties -ErrorAction Stop 
                        If (-not $Obj.Enabled) { $Obj | Get-ADReplicationAttributeMetadata -Properties UserAccountControl @Param -PipelineVariable UAC -ErrorAction SilentlyContinue | Select-Object $Select }
                        Else {
                            $Msg = "Object is not currently disabled!"
                            Write-Warning "[$i] $Msg"
                        }
                    }
                    Computer {
                        $Msg = " Looking for Active Directory computer"
                        Write-Verbose "[$i] $Msg"
                        Write-Progress -Activity $Activity -CurrentOperation $i -Status "ObjectClass: $ObjectClass" -PercentComplete ($Current / $Total * 100)
                        $Obj = Get-ADComputer -Identity $i -Properties $Properties @Param -ErrorAction Stop
                        If (-not $Obj.Enabled) { $Obj | Get-ADReplicationAttributeMetadata -Properties UserAccountControl @Param -PipelineVariable UAC -ErrorAction SilentlyContinue | Select-Object $Select }
                        Else {
                            $Msg = "Object is not currently disabled!"
                            Write-Warning "[$i] $Msg"
                        }
                    }
                    ServiceAccount {
                        $Msg = "Looking for Active Directory service account"
                        Write-Verbose "[$i] $Msg"
                        Write-Progress -Activity $Activity -CurrentOperation $i -Status "ObjectClass: $ObjectClass" -PercentComplete ($Current / $Total * 100)
                        $Obj = Get-ADServiceAccount -Identity $i -Properties $Properties @Param 
                        If (-not $Obj.Enabled) { $Obj | Get-ADReplicationAttributeMetadata -Properties UserAccountControl @Param -PipelineVariable UAC -ErrorAction SilentlyContinue | Select-Object $Select }
                        Else {
                            $Msg = "Object is not currently disabled!"
                            Write-Warning "[$i] $Msg"
                        }
                    }
                    Any {
                        $Msg = "Looking for Active Directory object"
                        Write-Verbose "[$i] $Msg"
                        Write-Progress -Activity $Activity -CurrentOperation $i -Status "ObjectClass: $ObjectClass" -PercentComplete ($Current / $Total * 100)
                        If (($i -match "^CN=") -or ($i -as [guid])) {
                            $Obj = Get-ADObject -Identity $i -Properties $Properties @Param -ErrorAction Stop 
                            If ($Obj.UserAccountControl -band 0x2) { $Obj | Get-ADReplicationAttributeMetadata -Properties UserAccountControl @Param -PipelineVariable UAC -ErrorAction SilentlyContinue | Select-Object $Select }
                            Else {
                                $Msg = "Object is not currently disabled!"
                                Write-Warning "[$i] $Msg"
                            }
                        }
                        Else {
                            $Objects = Get-ADObject -Filter "(Name -eq '$i') -or (SAMAccountName -eq '$i')" -Properties $Properties @Param -ErrorAction Stop
                            Foreach ($Obj in $Objects) {
                                If ($Obj.UserAccountControl -band 0x2) { $Obj | Get-ADReplicationAttributeMetadata -Properties UserAccountControl @Param -PipelineVariable UAC -ErrorAction SilentlyContinue | Select-Object $Select }
                                Else {
                                    $Msg = "Object is not currently disabled!"
                                    Write-Warning "[$i] $Msg"
                                }
                            }
                        }
                    }
                }
            }
            Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                $Msg = "Error!"
                If ($ObjectClass -ne "Any") { $Msg += " Consider setting ObjectClass to Any." }
                If ($_.Exception.Message) { $Msg += " $($_.Exception.Message)" }
                Write-Warning "[$i] $Msg"
            }
            Catch {
                $Msg = "Error!"
                If ($_.Exception.Message) { $Msg += " $($_.Exception.Message)" }
                Write-Warning "[$i] $Msg"
            }
        }
    }
    End {
        $Null = Write-Progress * -Completed
    }
} #end function