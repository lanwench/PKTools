#requires -Version 4
Function Convert-PKLastLogon {
<#
.SYNOPSIS
    Converts an Active Directory object's LastLogonTimestamp attribute value to human-readable datetime format

.DESCRIPTION
    Converts an Active Directory object's LastLogonTimestamp attribute value to human-readable datetime format
    Accepts pipeline input
    Outputs a DateTime object
    
.NOTES
    Name    : Function_Convert-PKLastLogon.ps1
    Author  : Paula Kingsley
    Created : 2020-06-23
    Version : 01.00.0000   
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2020-06-23 - Updated

.LINK
    https://techcommunity.microsoft.com/t5/ask-the-directory-services-team/8220-the-lastlogontimestamp-attribute-8221-8211-8220-what-it-was/ba-p/396204

.PARAMETER LastLogonTimestamp
    LastLogonTimestamp attribute value, as from Get-ADUser/Get-ADObject

.EXAMPLE
    PS C:\> $ Get-ADUser foobar -Properties LastLogonTimestamp | Convert-PKLastLogon -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                Value              
        ---                -----              
        Verbose            True               
        LastLogonTimestamp 0                  
        PipelineInput      True               
        ScriptName         Convert-PKLastLogon
        ScriptVersion      1.0.0              

        WARNING: It is important to note that the intended purpose of the lastLogontimeStamp attribute to help identify inactive computer and user accounts. 
        The lastLogon attribute is not designed to provide real time logon information. With default settings in place the lastLogontimeStamp will be 9-14 days behind the current date.
        See https://techcommunity.microsoft.com/t5/ask-the-directory-services-team/8220-the-lastlogontimestamp-attribute-8221-8211-8220-what-it-was/ba-p/396204

        VERBOSE: Time zone for all displayed dates is (UTC-08:00) Pacific Time (US & Canada)
        VERBOSE: [132369683195216552] Converting LastLogonTimestamp value to human-readable form

        Thursday, June 18, 2020 3:38:39 PM

.EXAMPLE
    PS C:\> Get-ADUser -SearchBase "OU=US-FTE,DC=domain,DC=com" -Properties EmailAddress,LastLogonTimestamp -Filter "Enabled -eq 'TRUE'" | 
        Select-Object Name,
        GivenName,
        Surname,
        SamAccountName,
        Enabled,
        @{N="LastLogonDate";E={$_.LastLogonTimestamp | Convert-PKLastLogon}},
        EmailAddress,
        UserPrincipalName,
        SID,
        DistinguishedName -OutVariable UserInfo    

            Name              : Wang, Deborah
            GivenName         : Deborah
            Surname           : Wang
            SamAccountName    : dwang
            Enabled           : True
            LastLogonDate     : 2020-06-19 6:52:42 PM
            EmailAddress      : Deborah.Wang@corp.net
            UserPrincipalName : Deborah.Wang@corp.net
            SID               : S-1-5-21-1606980848-223176313-839522115-3992563
            DistinguishedName : CN=Deborah Wang,OU=FTE,DC=domain,DC=local

            Name              : Bradley, Kamisha
            GivenName         : Kamisha
            Surname           : Bradley
            SamAccountName    : kbradley
            Enabled           : True
            LastLogonDate     : 2020-06-11 7:36:01 PM
            EmailAddress      : Kamisha.Bradley@corp.net
            UserPrincipalName : Kamisha.Bradley@corp.net
            SID               : S-1-5-21-1606980848-223176313-839522115-4099534
            DistinguishedName : CN=Kamisha Bradley,OU=FTE,DC=domain,DC=local

            Name              : Cruz, Ana
            GivenName         : Ana
            Surname           : Cruz
            SamAccountName    : acruz
            Enabled           : True
            LastLogonDate     : 2020-06-11 8:40:10 PM
            EmailAddress      : Ana.Cruz-Burton@corp.net
            UserPrincipalName : Ana.Cruz-Burton@corp.net
            SID               : S-1-5-21-1606980848-223166313-839522115-3762255
            DistinguishedName : CN=Ana Cruz,OU=FTE,DC=domain,DC=local


#>
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "LastLogonTimestamp attribute value, as from Get-ADUser/Get-ADObject"
    )]
    [int64]$LastLogonTimestamp

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"


$Disclaimer = @" 
It is important to note that the intended purpose of the lastLogontimeStamp attribute to help identify inactive computer and user accounts. 
The lastLogon attribute is not designed to provide real time logon information. With default settings in place the lastLogontimeStamp will be 9-14 days behind the current date.
See https://docs.microsoft.com/en-us/archive/blogs/askds/the-lastlogontimestamp-attribute-what-it-was-designed-for-and-how-it-works

"@

    Write-warning $Disclaimer

    $Timezone = ([System.TimeZoneInfo]::Local).DisplayName
    $Msg = "Time zone for all displayed dates is $TimeZone"
    Write-Verbose $Msg

    
}
Process {
    
    

    Foreach ($Logon in $LastLogonTimeStamp) {
        
        $Msg = "[$Logon] Converting LastLogonTimestamp value to [datetime]"
        Write-Verbose $Msg

        Try {
            If ($Results = [DateTime]::FromFileTimeutc($Logon)) {
                $Results
            }
            Else {
                $Msg = "[$Logon] Unable to convert input object"
                Write-Error $Msg
            }
        }
        Catch {}

    } #end foreach
}
} #end
