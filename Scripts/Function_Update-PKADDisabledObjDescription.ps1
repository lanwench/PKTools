#requires -version 4
Function Update-PKADDisabledObjDescription {
<#
.SYNOPSIS
    Updates the description field of disabled Active Directory objects with the object's disabled and last logon dates.

.DESCRIPTION
    The Update-PKADDisabledObjDescription function updates the description field of one or more disabled Active Directory objects with the object's disabled and lastlogon dates. 
    It accepts one or more AD objects or SAMAccountName strings as input and performs the specified action on each object. 
    The action can be to overwrite the current description, prepend the new description to the current one, or skip objects where the description already exists.

.NOTES
    Name    : Function_Update-PKADDisabledObjDescription.ps1
    Created : 2024-05-09
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2024-05-09 - Created script

.PARAMETER Name
    Specifies one or more AD objects or SAMAccountName strings. This parameter is mandatory.

.PARAMETER Action
    Specifies the action to perform on each object. The valid values are "Overwrite", "Prepend", and "SkipExisting". This parameter is mandatory.

.PARAMETER Server
    Specifies the name of the domain controller to connect to. If not specified, the current user's DNS domain is used.

.PARAMETER Credential
    Specifies alternate credentials to use for the operation.

.EXAMPLE
    PS C:\> Update-PKADDisabledObjDescription -Name User1 -Action SkipExisting
    Updates the AD object with SAMAccountName User1 to add the disabled and last logon dates to the Description field if there is no current Description value.

.EXAMPLE
    PS C:\> Update-PKADDisabledObjDescription -Name $Usernames -Action Prepend
    Updates the AD objects with the SAMAccountNames in an array by prepending the disabled and last logon dates to the current Description.

.EXAMPLE
    PS C:\> Get-ADUser -Filter * | Update-PKADDisabledObjDescription -Action Overwrite
    Updates the description field of all AD users by overwriting the current description with the disabled and last logon dates.


        
#>
    [CmdletBinding(SupportsShouldProcess,ConfirmImpact="High")]
Param(
    [Parameter(
        Position = 0,
        HelpMessage = "One or more Active Directory objects or SAMAccountName strings",
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        Mandatory
    )][Object[]]$Name,

    [Parameter(
        Mandatory,
        HelpMessage = "Overwrite current Description, prepend new Description to current, or skip objects where where Description already exists"
    )]
    [ValidateSet("Overwrite","Prepend","SkipExisting")]
    [string]$Action,

    [Parameter(
        HelpMessage = "Name of domain controller (default is current user's DNS domain)"
    )]
    [ValidateNotNullOrEmpty()]
    [Object[]]$Server = $Env:USERDNSDOMAIN,

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
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
    If (-not (Get-Module ActiveDirectory -ListAvailable -ErrorAction Stop)){
        Write-Error "ActiveDirectory module not present!"
        Break
    }

    $Param = @{Verbose = $False;ErrorAction = "Stop"}
    If ($PSBoundParameters.ContainsKey($Credential)) {
        $Param['Credential'] = $Credential
    }
    If (-not ($PSBoundParameters.ContainsKey($Server))) {
        $Param['Server'] = (Get-ADDomainController @Param).Hostname
    }

    $Properties = "Description,LastLogonTimestamp,UserAccountControl" -split(",")
    $Activity = "Update Active Directory object Description field with disabled date & last logon date (action: $Action)"
    Write-Verbose "[BEGIN: $ScriptName] $Activity"

}
Process {

    $Total = $Name.Count
    $Current = 0
    Foreach ($N in $Name) {
        $Current ++
        Try {
            $ADObject = $DisabledDate = $Null
            $Msg = "Getting Active Directory object"
            Write-Verbose "[$N] $Msg"
            Write-Progress -Activity $Activity -Status $Msg -CurrentOperation $N -PercentComplete($Current/$Total*100)

            If ($N -is [string]) {
                If (($N -match "^CN=") -or ($N -as [guid])) {
                    $ADObject = Get-ADObject -Identity $N -Properties $Properties @Param
                }
                Else {
                    $ADObject = Get-ADObject -Filter "(SAMAccountName -eq '$N')" -Properties $Properties @Param -ErrorAction Stop
                }
            }
            Elseif ($N -is [Microsoft.ActiveDirectory.Management.ADObject]) {
                $ADObject = $N | Get-ADObject -Properties * @Param 
            }

            If ($ADObject) {
                $Msg = "Found $($ADObject.ObjectClass) object"
                
                If (($ADObject.userAccountControl -band 0x0002) -ne 0) {
                    $DisabledDate = (Get-Date (($ADObject  | Get-ADReplicationAttributeMetadata -Properties UserAccountControl @Param).LastOriginatingChangeTime) -format u)
                    $Msg += ", disabled on $DisabledDate"
                    Write-Verbose "[$N] $Msg"

                    If ($ADObject.Description.length -gt 0 -and $Action -eq "SkipExisting") {
                        $Msg = "Description field not empty; -SkipExisting specified"
                        Write-Verbose "[$N] $Msg"
                    }
                    Else {
                        $NewDescription = "~Disabled $DisabledDate (last login $(Get-Date $([datetime]::FromFileTime($ADObject.LastLogonTimestamp)) -Format u))"
                        Switch ($Action) {
                            Overwrite {
                                $ConfirmMsg = "Setting object Description to $NewDescription"
                            }
                            Prepend {
                                If ($ADObject.Description.length -gt 0) {$NewDescription += " - $($ADObject.Description)"}
                                $ConfirmMsg = "Setting object Description to $NewDescription"
                            }
                        }
                        
                        Write-Verbose "[$N] $ConfirmMsg"
                        If ($PSCmdlet.ShouldProcess($N,$ConfirmMsg) ) {
                            $ADObject | Set-ADObject -Description $NewDescription -Confirm:$False @Param -Passthru
                        }
                        Else {
                            $Msg = "Operation cancelled by user"
                            Write-Verbose "[$N] $Msg"
                        }
                    }
                } #end if object disabled
                Else {
                    $Msg += " but object is not disabled!"
                    Write-Warning "[$N] $Msg"
                }
            } # end if object
            Else {
                $Msg = "No matching Active Directory object found"
                Write-Warning "[$N] $Msg"
            }
        }
        Catch {
            $Msg = "Something's wrong! $($_.Exception.Message)"
            Write-Warning "[$N] $Msg"
        }

    } #end foreach
}
End {
    Write-Verbose "[END: $ScriptName]"
    $Null = Write-Progress * -Completed
}
} #end function

