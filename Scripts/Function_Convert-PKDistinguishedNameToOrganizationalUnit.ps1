#requires -Version 3
Function Convert-PKDistinguishedNameToOrganizationalUnit {
<#
.SYNOPSIS
    Converts an Active Directory object's DistinguishedName to its parent Organizational Unit or Container, displayed in DN or CanonicalName format
   
.DESCRIPTION
    Converts an Active Directory object's DistinguishedName to its parent Organizational Unit or Container, displayed in DN or CanonicalName format
    Uses regular expressions to validate and parse DN strings (whether in organizational units or built-in containers)
    Returns a string
    Accepts pipeline input

.NOTES
    Name    : Function_Convert-PKDistinguishedNameToOrganizationalUnit.ps1 
    Created : 2020-02-20
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.00.0000 - 2020-02-20 - Created script

.PARAMETER DistinguishedName
    One or more Active Directory object DistinguishedNames (e.g., 'CN=computer123,OU=servers,DC=domain,DC=local'

.PARAMETER ConvertToCanonicalName
    Display output as CanonicalName instead of default DistinguishedName format, via string manipulation (e.g., 'domain.local/servers/computer123'

.EXAMPLE
    PS C:\> Get-ADGroup -Filter "Name -like 'ops-*'" | Convert-PKDistinguishedNametoOrganizationalUnit

        VERBOSE: PSBoundParameters: 
	
        Key                    Value                                          
        ---                    -----                                          
        DistinguishedName                                                     
        ConvertToCanonicalName False                                          
        PipelineInput          True                                           
        ScriptName             Convert-PKDistinguishedNametoOrganizationalUnit
        ScriptVersion          1.0.0                                          


        VERBOSE: [CN=ops-team-core,OU=Operations,OU=security,OU=Groups,DC=domain,DC=local]
        OU=Operations,OU=security,OU=Groups,DC=domain,DC=local
        VERBOSE: [CN=ops-support,OU=Distribution,OU=Groups,DC=domain,DC=local]
        OU=Distribution,OU=Groups,DC=domain,DC=local

.EXAMPLE
    PS C:\> Get-ADGroup -Filter "Name -like 'ops-*'"| Convert-PKDistinguishedNametoOrganizationalUnit -ConvertToCanonicalName

        domain.local/Groups/Security/Operations
        domain.local/Groups/Distribution/AquiredCo

.EXAMPLE
    PS C:\> Convert-GNDistinguishedNameToOrganizationalUnit -DistinguishedName "CN=UM Management,CN=Microsoft Exchange Security Groups,DC=domain,DC=local"

        CN=Microsoft Exchange Security Groups,DC=domain,DC=local

.EXAMPLE
    PS C:\> Convert-PKDistinguishedNameToOrganizationalUnit -DistinguishedName "OU=hellokitty,DC=domain,DC=local",jbloggs -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                    Value                                          
        ---                    -----                                          
        DistinguishedName      {OU=hellokitty,DC=domain,DC=local, jbloggs}             
        ConvertToCanonicalName False                                          
        PipelineInput          False                                          
        ScriptName             Convert-PKDistinguishedNametoOrganizationalUnit
        ScriptVersion          1.0.0                                          

        Convert-PKDistinguishedNameToOrganizationalUnit : [OU=hellokitty,DC=domain,DC=local] does not appear to be a valid object DistinguishedName
        At line:1 char:1
        + Convert-PKDistinguishedNameToOrganizationalUnit -DistinguishedName "O ...
        + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            + CategoryInfo          : NotSpecified: (:) [Write-Error], WriteErrorException
            + FullyQualifiedErrorId : Microsoft.PowerShell.Commands.WriteErrorException,Convert-PKDistinguishedNametoOrganizationalUnit

        Convert-PKDistinguishedNameToOrganizationalUnit : [jbloggs] does not appear to be a valid object DistinguishedName
        At line:1 char:1
        + Convert-PKDistinguishedNameToOrganizationalUnit -DistinguishedName "O ...
        + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            + CategoryInfo          : NotSpecified: (:) [Write-Error], WriteErrorException
            + FullyQualifiedErrorId : Microsoft.PowerShell.Commands.WriteErrorException,Convert-PKDistinguishedNametoOrganizationalUnit
 


#>
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more Active Directory object DistinguishedNames (e.g., 'CN=computer123,OU=servers,DC=domain,DC=local'"
    )]
    [string[]]$DistinguishedName,

    [Parameter(
        HelpMessage = "Display output as CanonicalName instead of default DistinguishedName format, via string manipulation (e.g., 'domain.local/servers/computer123'"
    )]
    [switch]$ConvertToCanonicalName

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    # Display our parameters
    $CurrentParams = $PSBoundParameters

    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #endregion Show parameters

    #region Function

    If ($ConvertToCanonicalName.IsPesent) {

        Function Get-CanonicalName ([string[]]$DistinguishedName) {
            # https://gallery.technet.microsoft.com/scriptcenter/Get-CanonicalName-Convert-a2aa82e5    
            Foreach ($dn in $DistinguishedName) {      
            
                ## Split the dn string up into its constituent parts 
                $d = $dn.Split(',') 

                ## get parts excluding the parts relevant to the FQDN and trim off the dn syntax 
                $arr = (@(($d | Where-Object { $_ -notmatch 'DC=' }) | ForEach-Object { $_.Substring(3) }))  
            
                ## Flip the order of the array. 
                [array]::Reverse($arr)  
 
                ## Create and return the string representation in canonical name format of the supplied DN 
                "{0}/{1}" -f  (($d | Where-Object { $_ -match 'dc=' } | ForEach-Object { $_.Replace('DC=','') }) -join '.'), ($arr -join '/') 
            } 
        }
    }

    #endregion Function

    #region And now we have two problems
        
        # A little excessive as we don't actually need named capture groups, but, why not?
        $Regex = "^(?:(?<cn>CN=(?<name>[^,]*)),)?(?:(?<path>(?:(?:CN|OU)=[^,]+,?)+),)?(?<domain>(?:DC=[^,]+,?)+)$"

    #endregion And now we have two problems

}
Process {

    Foreach ($DN in $DistinguishedName) {
        
        If ($DN -match "^CN=,(OU|CN)=") {
            Write-Verbose "[$DN]"
            $Output = $DN -Replace '^.+?(?<!\\),',''
            If ($ConvertToCanonicalName.IsPresent) {
                Get-CanonicalName -DistinguishedName $Output
            }
            Else {
                Write-Output $Output
            }
        }
        Else {
            $Msg = "[$DN] does not appear to be a valid object DistinguishedName"
            Write-Error $Msg
        }
    }
}
End {}
} #end Convert-PKDistinguishedNameToOrganizationalUnit

$Null = New-Alias -Name Convert-PKDNtoOU -Value Convert-PKDistinguishedNameToOrganizationalUnit -Force -Confirm:$False