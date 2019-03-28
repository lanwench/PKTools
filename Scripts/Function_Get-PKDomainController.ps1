#requires -Version 3
Function Get-PKADDomainController {
<#
.SYNOPSIS
    Returns a domain controller for the current computer site or a named site in the current or named domain
    
.DESCRIPTION
   Returns a domain controller for the current computer site or a named site in the current or named domain
   If the first match fails, tries adjacent sites; if that fails, returns the first DC in the domain
   Uses .NET System.DirectoryServices.ActiveDirectory.DirectoryContext (no ActiveDirectory module needed)
   Returns a DomainController object (System.DirectoryServices.ActiveDirectory.DirectoryServer)
   
.NOTES
    Name    : Function_Get-PKADDomainController.ps1
    Created : 2017-09-15
    Author  : Paula Kingsley
    Version : 01.01.0000
    History:

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK *

        v01.00.0000 - 2017-09-15 - Created script
        v01.01.0000 - 2019-03-26 - General cosmetic udates
        
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
        ADDomain      domain.local
        ScriptName    Get-PKDomainController 
        ScriptVersion 1.0.0                  

        VERBOSE: Get domain details
        VERBOSE: Domain 'domain.local' is running in Windows2008R2Domain mode in forest 'domain.local'
        VERBOSE: Get site
        VERBOSE: Current site is 'Sacramento'
        VERBOSE: Get domain controller in site
        VERBOSE: Found domain controller 'DC04.domain.local'


        Forest                     : domain.local
        CurrentTime                : 2017-09-15 20:51:56
        HighestCommittedUsn        : 588463910
        OSVersion                  : Windows Server 2008 R2 Enterprise
        Roles                      : {SchemaRole, NamingRole, PdcRole, RidRole...}
        Domain                     : domain.local
        IPAddress                  : 10.8.142.103
        SiteName                   : Sacramento
        SyncFromAllServersCallback : 
        InboundConnections         : {6ba69348-a596-4f09-b1a4-652e44b6fea0, EVLDC00, GFDC01, SCDC03}
        OutboundConnections        : {198b36d4-2bbb-4e76-9079-e9cc53c9196f}
        Name                       : DC04.domain.local
        Partitions                 : {DC=domain,DC=local, CN=Configuration,DC=domain,DC=local, 
                                     CN=Schema,CN=Configuration,DC=domain,DC=local, 
                                     DC=DomainDnsZones,DC=domain,DC=local...}



#>
[CmdletBinding()]
Param (
    [Parameter(
        ValueFromPipeline = $True,
        HelpMessage = "Active Directory object, name or FQDN (default is current user)"
    )]
    [ValidateNotNullOrEmpty()]
    [object]$ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain(),
    
    [Parameter(
        ParameterSetName = "BySite",
        HelpMessage = "AD site (default is all)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$ADSite, # = [System.DirectoryServices.ActiveDirectory.ActiveDirectorysite]::GetComputerSite(),

    [Parameter(
        HelpMessage ="Suppress all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)
Begin {

    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"

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

    #region Functions

    # Get forest
    Function GetDomain {
        [CmdletBinding()]
        Param($DomainName,$Credential)
        Try {
            $ContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]'Domain' 
            If ($Credential) {
                $UserName = $Credential.Username
                $Password = $Credential.GetNetworkCredential().Password
                If ($Username -match "\\") {$Username = $Username.Split("\")[1]}
                Elseif ($Username -match "@") {$Username = $Username.Split("@")[0]}
                $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext($ContextType,$DomainName,$UserName,$Password)
            }
            Else {
                $Context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext($ContextType,$DomainName)
            }
            [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)
        }
        Catch {}
    }

    # Get forest
    Function GetForest {
        [CmdletBinding()]
        Param($ADDomain)
        Try {
            $ContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]'Forest' 
            New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext($ContextType,$ADDomain.Forest) 
            }
        Catch {}
    }

    # Get sites
    Function GetSite {
        [CmdletBinding()]
        Param($Forest,$SiteName)
        Try {
            $AllSites = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($Forest).sites
            If ($SiteName) {
                $AllSites | Where-Object {$_.Name -eq $SiteName}
            }
            Else {$AllSites}
        }
        Catch {}
    }

    #endregion Functions

    # Get domain
    If ($ADDomain.GetType().Name -eq "String") {
        $Msg = "Get Active Directory domain"
        Write-Verbose "[Prerequisites] $Msg"
        
        # https://social.technet.microsoft.com/Forums/en-US/c771e6bf-5134-4d25-8cef-8b4f9a136627/binding-to-ad?forum=ITCG
        
        $ADDomain = GetDomain -DomainName $ADDomain -Credential $Credential
    }    
    
    # Get forest
    $ADForest= GetForest -ADDomain $ADDomain


    # Get sites
    $ADSites = GetSite -Forest $ADForest



    #region Splats

    # Splat for Write-Progress
    $Activity = "Get Active Directory domain controllers using .NET"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
    }

    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}

}

Process {
    
    [switch]$Continue = $False

    $Msg = "Get Active Directory domain"
    Write-Verbose "[$ADDomain] $Msg"
    $CurrentOperation = $Msg
    Write-Progress -Activity $Activity -CurrentOperation $CurrentOperation

    Try {
        $DomainType = [System.DirectoryServices.ActiveDirectory.DirectoryContexttype]"Domain"
        $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext($DomainType, $ADDomain)
        $DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)
            
        $Msg = "Domain '$($DomainObj.Name)' is running in $($DomainObj.DomainMode) mode in forest '$($DomainObj.Forest)'"
        Write-Verbose "[$ADDomain] $Msg"
        $Continue= $True

    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = ($_.Exception.Message.Split(':')[1]).Trim().Replace('"','').Replace('.','')) {$Msg += "`n$ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR  : [$ADDomain] $Msg")
        Break
    }
    
    If ($Continue.IsPresent) {
    
        # reset flag
        $Continue = $False    

        $Msg = "Get Active Directory site"
        Write-Verbose "[$ADDomain] $Msg"
        $CurrentOperation = $Msg
        Write-Progress -Activity $Activity -CurrentOperation $CurrentOperation

        Try {
            $SiteContext = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $DomainObj.Forest)
            If (-not ($SiteObj = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($SiteContext).sites | Where-Object {$_.Name -eq $ADSite})){
                $Msg = "Invalid AD site '$ADSite"
                $Host.UI.WriteErrorLine("ERROR  : [$ADDomain] $Msg")
            }
            Else {
                $SiteName = $SiteObj.Name
                $Msg = "Current site is '$SiteName'"
                Write-Verbose "[$ADDomain] $Msg"
                $Continue = $True
            }
        }
        Catch {
            $Msg = "Operation failed for site '$ADSite'"
            If ($ErrorDetails = ($_.Exception.Message.Split(':')[1]).Trim().Replace('"','').Replace('.','')) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR  : [$ADDomain] $Msg")
               
        }
    }
    
    If ($Continue.IsPresent) {
    
        # reset flag
        $Continue = $False    

        $Msg = "Get Get Active Directory domain controller(s) in site"
        Write-Verbose "[$ADDomain] $Msg"
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
                        Write-Warning "[$ADDomain] $Msg"
                        $DC = $DomainObj.DomainControllers[0]
                    }
                } #end foreach adjacent site
            } # end get DC in site

            If ($DC) {
                $Msg = "Found domain controller '$($DC.Name)'"
                Write-Verbose "[$ADDomain] $Msg"
                Write-Output $DC
            }
            Else {
                $Msg = "No domain controller found"
                $Host.UI.WriteErrorLine("ERROR  : [$ADDomain] $Msg")
            }       
        }
        Catch {
            $Msg = "Domain controller lookup failed"
            If ($ErrorDetails = ($_.Exception.Message.Split(':')[1]).Trim().Replace('"','').Replace('.','')) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR  : [$ADDomain] $Msg")
        }
    }
    
}
End {
    
    Write-Progress -Activity $Activity -Completed
}
} #end Get-PKADDomainController




