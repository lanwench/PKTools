#requires -Version 3
Function Get-PKADDomainController {
<#

.Synopsis
    Returns a domain controller for the current computer site or a named site in the current or named domain
    
.DESCRIPTION
   Returns a domain controller for the current computer site or a named site in the current or named domain
   If the first match fails, tries adjacent sites; if that fails, returns the first DC in the domain
   Uses .NET System.DirectoryServices.ActiveDirectory.DirectoryContext (no ActiveDirectory module needed)
   Returns a DomainController object (System.DirectoryServices.ActiveDirectory.DirectoryServer)
   
.NOTES
    Name    : Function_Get-PKADDomainController.ps1
    Created : 2017-09-15
    Version : 01.00.0000
    Author  : Paula Kingsley
    
    History:

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK *

        v01.00.0000 - 2017-09-15 - Created script
        
.EXAMPLE
    PS C:\> Get-PKDomainController -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                  
        ---           -----                  
        Verbose       True                   
        ADDomain      domain.local
        ADSite        Las Vegas             
        ScriptName    Get-PKDomainController 
        ScriptVersion 1.0.0                  

        VERBOSE: Get domain details
        VERBOSE: Domain 'domain.local' is running in Windows2008R2Domain mode in forest 'domain.local'
        VERBOSE: Get computer site
        VERBOSE: Current site is 'Las Vegas'
        VERBOSE: Get domain controllers in site
        VERBOSE: Found domain controller 'DC08.domain.local'


        Forest                     : domain.local
        CurrentTime                : 2017-09-15 20:45:37
        HighestCommittedUsn        : 162461849
        OSVersion                  : Windows Server 2008 R2 Enterprise
        Roles                      : {}
        Domain                     : domain.local
        IPAddress                  : 10.15.144.11
        SiteName                   : Las Vegas
        SyncFromAllServersCallback : 
        InboundConnections         : {026fc520-6d52-4d33-9c61-b2b8950e270a, 1bde646a-2f1f-4f66-b8b5-7ac544f4637e, 
                                     1f3b6863-d5ab-4a86-8b9a-6ec78fdbfe5e, 6ed28c8c-ca6c-4556-900c-b137b5b94ed2...}
        OutboundConnections        : {6c98b0bc-83fe-4ce7-85de-22bc4e46d442}
        Name                       : DC08.domain.local
        Partitions                 : {DC=domain,DC=local, CN=Configuration,DC=domain,DC=local, 
                                     CN=Schema,CN=Configuration,DC=domain,DC=local, 
                                     DC=DomainDnsZones,DC=domain,DC=local...}
.EXAMPLE
    PS C:\> Get-PKDomainController -ADSite Sacramento -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                  
        ---           -----                  
        ADSite        Sacramento             
        Verbose       True                   
        ADDomain      gracenote.gracenote.com
        ScriptName    Get-PKDomainController 
        ScriptVersion 1.0.0                  



        VERBOSE: Get domain details
        VERBOSE: Domain 'gracenote.gracenote.com' is running in Windows2008R2Domain mode in forest 'gracenote.gracenote.com'
        VERBOSE: Get site
        VERBOSE: Current site is 'Sacramento'
        VERBOSE: Get domain controller in site
        VERBOSE: Found domain controller 'SACDC01.gracenote.gracenote.com'


        Forest                     : gracenote.gracenote.com
        CurrentTime                : 2017-09-15 20:51:56
        HighestCommittedUsn        : 588463910
        OSVersion                  : Windows Server 2008 R2 Enterprise
        Roles                      : {SchemaRole, NamingRole, PdcRole, RidRole...}
        Domain                     : gracenote.gracenote.com
        IPAddress                  : 10.8.142.103
        SiteName                   : Sacramento
        SyncFromAllServersCallback : 
        InboundConnections         : {6ba69348-a596-4f09-b1a4-652e44b6fea0, EVLDC00, GFDC01, SCDC03}
        OutboundConnections        : {198b36d4-2bbb-4e76-9079-e9cc53c9196f}
        Name                       : SACDC01.gracenote.gracenote.com
        Partitions                 : {DC=gracenote,DC=gracenote,DC=com, CN=Configuration,DC=gracenote,DC=gracenote,DC=com, 
                                     CN=Schema,CN=Configuration,DC=gracenote,DC=gracenote,DC=com, 
                                     DC=DomainDnsZones,DC=gracenote,DC=gracenote,DC=com...}



#>
[CmdletBinding()]
Param (
    [Parameter(
        Mandatory = $False,
        HelpMessage = "AD domain name FQDN (default is current computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain(),

    [Parameter(
        Mandatory = $False,
        HelpMessage = "AD site (default is current computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorysite]::GetComputerSite()

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

    # Preference
    $ErrorActionPreference = "Stop"

}

Process {
    
    $Msg = "Get domain details"
    Write-Verbose $Msg

    $Activity = "Get domain controller in site"
    $CurrentOperation = $Msg
    Write-Progress -Activity $Activity -CurrentOperation $CurrentOperation

    Try {
        $DomainType = [System.DirectoryServices.ActiveDirectory.DirectoryContexttype]"Domain"
        $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext($DomainType, $ADDomain)
        $DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)
            
        $Msg = "Domain '$($DomainObj.Name)' is running in $($DomainObj.DomainMode) mode in forest '$($DomainObj.Forest)'"
        Write-Verbose $Msg

    }
    Catch {
        $Msg = "Operation failed for domain '$ADDomain'"
        $ErrorDetails = ($_.Exception.Message.Split(':')[1]).Trim().Replace('"','').Replace('.','')
        $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
        Break
    }
        

    $Msg = "Get site"
    Write-Verbose $Msg
    $CurrentOperation = $Msg
    Write-Progress -Activity $Activity -CurrentOperation $CurrentOperation

    Try {
        $SiteContext = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $DomainObj.Forest)
        If (-not ($SiteObj = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($SiteContext).sites | Where-Object {$_.Name -eq $ADSite})){
            $Msg = "Invalid AD site '$ADSite"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
        Else {
            $SiteName = $SiteObj.Name
            $Msg = "Current site is '$SiteName'"
            Write-Verbose $Msg
        }
    }
    Catch {
        $Msg = "Operation failed for site '$ADSite'"
        $ErrorDetails = ($_.Exception.Message.Split(':')[1]).Trim().Replace('"','').Replace('.','')
        $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
        Break
            
    }
    
    $Msg = "Get domain controller in site"
    Write-Verbose $Msg
    $CurrentOperation = $Msg
    Write-Progress -Activity $Activity -CurrentOperation $CurrentOperation

    Try {
        If (-not ($DC = $DomainObj.DomainControllers | Where-Object {$_.SiteName -eq $SiteName} | Select -first 1)) {
            $Msg = "Failed to get domain controller for '$SiteName'; check connected sites"
            Write-Warning $Msg
            $SiteObj.AdjacentSites | Foreach-Object {
                $SiteName = $_.Name
                If (-not ($DC = $DomainObj.DomainControllers | Where-Object {$_.SiteName -eq $SiteName} | Select -first 1)) {
                    $Msg = "Failed to get domain controller for '$SiteName'; select first DC in domain"
                    Write-Warning $Msg
                    $DC = $DomainObj.DomainControllers[0]
                }
            } #end foreach adjacent site
        } # end get DC in site

        If ($DC) {
            $Msg = "Found domain controller '$($DC.Name)'"
            Write-Verbose $Msg
            Write-Output $DC
        }
        Else {
            $Msg = "No domain controller found"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
        }       
    }
    Catch {
        $Msg = "Domain controller lookup failed"
        $ErrorDetails = $_.Exception.Message
        $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
    }

    Write-Progress -Activity $Activity -Completed
}
} #end Get-PKADDomainController



