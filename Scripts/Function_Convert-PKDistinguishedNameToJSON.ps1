#Requires -version 4
Function Convert-PKDistinguishedNameToJSON {
<# 
.SYNOPSIS
    Uses Zachary Loeber's Get-ChildOUStructure to output the CanonicalName format of a container/OU as JSON 

.DESCRIPTION
    Uses Zachary Loeber's Get-ChildOUStructure to output the CanonicalName format of a container/OU as JSON 
    If no Searchbase provided, domain or server can be provided; defaults to root of current user's domain
    Accepts pipeline input
    Returns a JSON object

.NOTES        
    Name    : Function_Convert-PKDistinguishedNameToJSON.ps1
    Created : 2020-02-21
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-02-21 - Created script

.PARAMETER SearchBase              
    Starting Active Directory path/organizational unit (default is root of current user's domain)
        

.PARAMETER ADDomain             
    Active Directory domain name or FQDN (default is current user's domain)

.PARAMETER Server               
    Domain controller name or FQDN (default is first available)

.PARAMETER Depth
    Depth for JSON output (default is 20)

.PARAMETER Quiet
    Suppress non-verbose console output (outputs errors as errors, not formatted strings)


.EXAMPLE
    PS C:\> 

#> 

[CmdletBinding(
    DefaultParameterSetName = "Default",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    
    [Parameter(
        ParameterSetName = "Named",
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Active Directory path/organizational unit (if not provided, defaults to root of current user's domain)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,
    
    [Parameter(
        ParameterSetName = "Default",
        HelpMessage = "Active Directory domain name or FQDN (default is current user's domain)"
    )]
    [ValidateNotNullOrEmpty()]
    [string] $ADDomain,

    [Parameter(
        ParameterSetName = "Default",
        HelpMessage = "Domain controller name or FQDN (default is first available)"
    )]
    [ValidateNotNullOrEmpty()]
    [String] $Server,

    [Parameter(
        HelpMessage = "Depth for JSON output (default is 20)"
    )]
    [ValidateNotNullOrEmpty()]
    [int] $Depth = 20,

    [Parameter(
        HelpMessage = "Suppress non-verbose console output (outputs errors as errors, not formatted strings)"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch]$Quiet


)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences
    $ErrorActionPreference = "Stop"
    
    #region Functions
    # Mix, match, use, discard, whatever

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

    # Function to write an error as a string (no stacktrace), or an error, and options for prefix to string
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        If (-not $Quiet.IsPresent) {
            $Host.UI.WriteErrorLine("$Message")
        }
        Else {Write-Error "$Message"}
    }
    # Function to write a warning, with any error data, and options for prefix to string
    Function Write-MessageWarning {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Warning $Msg
    }

    # Function to write an error/warning, collecting error data
    Function Write-MessageVerbose {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Verbose $Msg
    }

    # Function to connect to AD with (named or default) AD domain, and (named or default) DC, in order of preference
    # Outputs global variables for domain object, dc object, and domain name string, and dc name string
    Function ConnectTo-AD {
        [CmdletBinding()]
        Param()
        $ErrorActionPreference = "Stop"

        # Splats
        $Param_GetDomain = @{}
        $Param_GetDomain = @{Identity = $Null}
        If ($CurrentParams.Credential) {$Param_GetDomain.Add("Credential",$Credential)}

        $Param_GetDC = @{}
        If ($CurrentParams.Credential) {$Param_GetDC.Add("Credential",$Credential)}

        # In order of preference based on parameters

        #region Get a named domain and get any domain controller (this is the easiest)

        If ($CurrentParams.ADDomain -and (-not $CurrentParams.Server)) {
        
            Try {
                $Msg = "Get Active Directory domain '$($CurrentParams.ADDomain)'"
                $Msg | Write-MessageVerbose -PrefixPrerequisites
                $Param_GetDomain.Identity = $CurrentParams.ADDomain
                $DomainObj = Get-ADDomain @Param_GetDomain

                Try {
                    $Msg = "Get first available domain controller in '$($DomainObj.NetBIOSName)'"
                    $Msg | Write-MessageVerbose -PrefixPrerequisites
                    $Param_GetDC.Add("DomainName",$DomainObj.DNSRoot)
                    $DCObj = Get-ADDomainController -Discover -ForceDiscover @Param_GetDC 
                }
                Catch {
                    $Msg = "Failed to get Active Directory domain controller"
                    $Msg | Write-MessageError -PrefixPrerequisites
                    Break
                }
            }
            Catch {
                $Msg = "Failed to get Active Directory domain"
                $Msg | Write-MessageError -PrefixPrerequisites
                Break
            }
        } #end if domain but not domain controller provided

        #endregion Get a named domain and get any domain controller

        #region ...Or get the domain and then the named domain controller (error out if DC isn't in that domain) 

        Elseif ($CurrentParams.ADDomain -and $CurrentParams.Server) {
            Try {
                $Msg = "Get Active Directory domain '$($CurrentParams.ADDomain)'"
                $Msg | Write-MessageVerbose -PrefixPrerequisites
                $Param_GetDomain.Identity = $CurrentParams.ADDomain
                $DomainObj = Get-ADDomain @Param_GetDomain

                Try {
                    $Msg = "Get domain controller '$($CurrentParams.Server)'"
                    $Msg | Write-MessageVerbose -PrefixPrerequisites
                    $Param_GetDC.Add("Identity",$CurrentParams.Server)
                    $Param_GetDC.Add("Server",$CurrentParams.Server)
                    $DCObj = Get-ADDomainController @Param_GetDC 
                    If ($DCObj.Domain -ne $DomainObj.DNSRoot)  {
                        $DCObj = $Null
                        $Msg = "Domain controller '$($CurrentParams.Server)' is not in domain '$($DomainObj.Domain)'"
                        $Msg | Write-MessageError -PrefixPrerequisites
                        Break
                    }   
                }
                Catch {
                    $Msg = "Failed to get Active Directory domain controller"
                    $Msg | Write-MessageError -PrefixPrerequisites
                    Break
                }
            }
            Catch {
                $Msg = "Failed to get Active Directory domain"
                $Msg | Write-MessageError -PrefixPrerequisites
                Break
            }

        } # end if domain and named DC

        #endregion Or get the domain and then the named domain controller (error out if DC isn't in that domain) 

        #region ...Or get a named domain controller and the domain it's in

        Elseif ($CurrentParams.Server -and (-not $CurrentParams.ADDomain)) {
            $Msg = "Get named domain controller and associated domain"
            $Msg | Write-Verbose  -PrefixPrerequisites
        
            Try {
                $Msg = "Get domain controller '$($CurrentParams.Server)'"
                $Msg | Write-MessageVerbose -PrefixPrerequisites
                $Param_GetDC.Add("Identity",$CurrentParams.Server)
                $Param_GetDC.Add("Server",$CurrentParams.Server)
                $DCObj = Get-ADDomainController @Param_GetDC
            
                Try {
                    $Msg = "Get Active Directory domain '$($DCObj.Domain)'"
                    $Msg | Write-MessageVerbose -PrefixPrerequisites
                    $Param_GetDomain.Identity = $DCObj.Domain
                    $DomainObj = Get-ADDomain @Param_GetDomain
                }
                Catch {
                    $Msg = "Failed to get Active Directory domain"
                    $Msg | Write-MessageError -PrefixPrerequisites 
                    Break
                }
            }
            Catch {
                $Msg = "Failed to get Active Directory domain controller"
                $Msg | Write-MessageError -PrefixPrerequisites
                Break
            }
        } # end if DC and no domain

        #endregion ...Or get a named domain controller and the domain it's in

        #region ...Or get the current user's domain and the first available domain controller

        Elseif ((-not $CurrentParams.Server) -and (-not $CurrentParams.ADDomain)) {
        
            Try {
                $Msg = "Get current user's Active Directory domain"
                $Msg | Write-MessageVerbose -PrefixPrerequisites
                $Param_GetDomain.Identity = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
                $DomainObj = Get-ADDomain @Param_GetDomain

                Try {
                    $Msg = "Get first available domain controller in '$($DomainObj.NetBIOSName)'"
                    $Msg | Write-MessageVerbose -PrefixPrerequisites
                    $Param_GetDC.Add("DomainName",$DomainObj.DNSRoot)
                    $DCObj = Get-ADDomainController -Discover -ForceDiscover @Param_GetDC 
                }
                Catch {
                    $Msg = "Failed to get Active Directory domain controller"
                    $Msg | Write-MessageError -PrefixPrerequisites
                    Break
                }
            }
            Catch {
                $Msg = "Failed to get Active Directory domain"
                $Msg | Write-MessageError -PrefixPrerequisites
                Break
            }

        } #end if no domain or DC provided

        #endregion ...Or get the current user's domain and the first available domain controller

        #region Now get or validate searchbase if we have a DC and domain, and output variables

        If ($DCObj -and $DomainObj) {
            
            If ($CurrentParams.SearchBase) {
                
                Try {
                    $Msg = "Validate searchbase"
                    $Msg | Write-MessageVerbose -PrefixPrerequisites
                    If ($LookupSearchBase = Get-ADObject -Identity $CurrentParams.SearchBase -Server $($DCObj.HostName) -ErrorAction SilentlyContinue| Where-Object {$_.ObjectClass -in @("OrganizationalUnit","builtinDomain","DomainDNS")}) {
                        If ($CurrentParams.SearchBase -notmatch $DomainObj.DistinguishedName) {
                            $Msg = "Searchbase '$($CurrentParams.SearchBase)' does not appear to match domain '$($DomainObj.DistinguishedName)'"
                            $Msg | Write-MessageError
                            Break
                        }
                        Else {
                            New-Variable -Name BaseDN -Scope Global -Value $LookupSearchBase.DistinguishedName -Force
                            $Msg = "Validated searchbase '$BaseDN'"                
                            $Msg | Write-MessageVerbose -PrefixPrerequisites
                        }
                    }
                    Else {
                        $Msg = "Failed to validate searchbase '$($CurrentParams.SearchBase)'"
                        $Msg | Write-MessageError -PrefixPrerequisites
                    }
                }
                Catch {
                    $Msg = "Failed to validate searchbase '$($CurrentParams.SearchBase)'"
                    $Msg | Write-MessageError -PrefixPrerequisites
                    Break
                }
            }
            Else {
                New-Variable -Name BaseDN -Scope Global -Value $DomainObj.DistinguishedName -Force                
                $Msg = "Setting searchbase to root of domain, '$BaseDN'"
                Write-Verbose $Msg
            }       
            
            If ($DCObj -and $DomainObj -and $BaseDN) {
                New-Variable -Name DomainObj -Scope Global -Value $DomainObj -Force
                New-Variable -Name DCObj -Scope Global -Value $DCObj -Force
                New-Variable -Name DC -Scope Global -Value $($DCObj.Hostname) -Force
                New-Variable -Name Domain -Scope Global -Value $($DomainObj.DNSRoot) -Force

                $Msg = "Successfully connected to domain controller '$DC' in domain '$Domain' with searchbase '$BaseDN'"
                $Msg | Write-MessageVerbose -PrefixPrerequisites
            }
        }

        #endregion Now get or validate searchbase if we have a DC and domain, and output variables
        
    } #end ConnectTo-AD

    # Convert DN to CN
    function Get-CanonicalName {
    Param (
        [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string[]]$DistinguishedName
    )     
        foreach ($dn in $DistinguishedName) {      
            $d = $dn.Split(',') ## Split the dn string up into it's constituent parts 
            $arr = (@(($d | Where-Object { $_ -notmatch 'DC=' }) | ForEach-Object { $_.Substring(3) }))  ## get parts excluding the parts relevant to the FQDN and trim off the dn syntax 
            [array]::Reverse($arr)  ## Flip the order of the array. 
 
            ## Create and return the string representation in canonical name format of the supplied DN 
            $("{0}/{1}" -f  (($d | Where-Object { $_ -match 'dc=' } | ForEach-Object { $_.Replace('DC=','') }) -join '.'), ($arr -join '/')).TrimEnd("/") 
        } 
    }

    # Function that actually does the stuff (Zachary Loeber)
    function Get-ChildOUStructure {
    <#
    .SYNOPSIS
        Create JSON exportable tree view of AD OU (or other) structures.

    .DESCRIPTION
        Create JSON exportable tree view of AD OU (or other) structures in Canonical Name format.

    .PARAMETER ouarray
        Array of OUs in CanonicalName format (ie. domain/ou1/ou2)

    .PARAMETER oubase
        Base of OU

    .EXAMPLE
        $OUs = @(Get-ADObject -Filter {(ObjectClass -eq "OrganizationalUnit")} -Properties CanonicalName).CanonicalName
        $test = $OUs | Get-ChildOUStructure | ConvertTo-Json -Depth 20

    .NOTES
        Author: Zachary Loeber
        Requires: Powershell 3.0, Lync
        Version History
        1.0.0 - 12/24/2014
            - Initial release
    .LINK
        https://github.com/zloeber/Powershell/blob/master/ActiveDirectory/Get-ChildOUStructure.ps1
    .LINK
        http://www.the-little-things.net
    #>
    [CmdletBinding()]
    param(
        [Parameter(
            Position=0, 
            ValueFromPipeline=$true, 
            Mandatory=$true, 
            HelpMessage='Array of OUs in CanonicalName format (ie. domain/ou1/ou2)'
        )]
        [string[]]$ouarray,

        [Parameter(Position=1, HelpMessage='Base of OU.')]
        [string]$oubase = ''
    )
    begin {
        $newarray = @()
        $base = ''
        $firstset = $false
        $ouarraylist = @()
    }
    process {
        $ouarraylist += $ouarray
    }
    end {
        $ouarraylist = $ouarraylist | Where {($_ -ne $null) -and ($_ -ne '')} | Select -Unique | Sort-Object
        if ($ouarraylist.count -gt 0) {
            $ouarraylist | Foreach {
                # $prioroupath = if ($oubase -ne '') {$oubase + '/' + $_} else {''}
                $firstelement = @($_ -split '/')[0]
                $regex = "`^`($firstelement`?`)"
                $tmp = $_ -replace $regex,'' -replace "^(\/?)",''

                if (-not $firstset) {
                    $base = $firstelement
                    $firstset = $true
                }
                else {
                    if (($base -ne $firstelement) -or ($tmp -eq '')) {
                        Write-Verbose "Processing Subtree for: $base"
                        $fulloupath = if ($oubase -ne '') {$oubase + '/' + $base} else {$base}
                        New-Object psobject -Property @{
                            'name' = $base
                            'path' = $fulloupath
                            'children' = if ($newarray.Count -gt 0) {,@(Get-ChildOUStructure -ouarray $newarray -oubase $fulloupath)} else {$null}
                        }
                        $base = $firstelement
                        $newarray = @()
                        $firstset = $false
                    }
                }
                if ($tmp -ne '') {
                    $newarray += $tmp
                }
            }
            Write-Verbose "Processing Subtree for: $base"
            $fulloupath = if ($oubase -ne '') {$oubase + '/' + $base} else {$base}
            New-Object psobject -Property @{
                'name' = $base
                'path' = $fulloupath
                'children' = if ($newarray.Count -gt 0) {,@(Get-ChildOUStructure -ouarray $newarray -oubase $fulloupath)} else {$null}
            }
        }
    }
    }

    #endregion Functions

    #region Prerequisites

    $Activity = "Prerequisites"

    # Make sure AD module is loaded
    # Use this only if not using a -Requires statement, which you may not want to do if this is part of a module 
    # containing non-AD-related functions too. Your call!

    $Msg = "Verify ActiveDirectory module"
    $Msg | Write-MessageVerbose -PrefixPrerequisites
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    Try {
        If ($Module = Get-Module -Name ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue -Verbose:$False) {
            $Msg = "Successfully located ActiveDirectory module version $($Module.Version.ToString())"
            $Msg | Write-MessageVerbose -PrefixPrerequisites
        }
        Else {
            $Msg = "Failed to find ActiveDirectory module in PSModule path"
            $Msg | Write-MessageError -PrefixPrerequisites
            Break
        }
    }
    Catch {
        $Msg = "Failed to find ActiveDirectory module"
        $Msg | Write-MessageError -PrefixPrerequisites
        Break
    }

    #region Prerequisites 

    $Msg = "Connect to Active Directory"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    ConnectTo-AD #-verbose

    If (-not ($DC -and $Domain -and $BaseDN)) {
        $Msg = "[Prerequisites] Failed to connect to Active Directory; please specify a valid domain name and/or domain controller name"
        $Msg | Write-MessageError
        Break
    }

    #endregion Prerequisites

    #endregion Prerequisites

    #region Splats

    # Splat for write-progress
    $Activity = "Create JSON output from Active Directory organizational unit tree (CanonicalName syntax)"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
    }

    #endregion Splats

    # Console output
    "[BEGIN: $Scriptname] $Activity" | Write-MessageInfo -FGColor Yellow -Title
    
} #end begin
Process {

    Foreach ($OU in $BaseDN) {
        
        $Param_WP.CurrentOperation = $OU
        Write-Progress @Param_WP
        $OU | Write-MessageVerbose
        
        Get-ChildOUStructure ($BaseDN | Get-CanonicalName) | ConvertTo-Json -Depth $Depth
                
    }
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    "[END $Scriptname] $Activity" | Write-MessageInfo -FGColor Yellow -Title
}

} # end Convert-PKDistinguishedNameToJSON

