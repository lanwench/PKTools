#requires -Version 3
Function Test-PKADConnection {
<#
.SYNOPSIS
    Tests connectivity to an Active Directory domain (and optional domain controller) without the ActiveDirectory module
    
.DESCRIPTION
    Tests connectivity to an Active Directory domain (and optional domain controller) without the ActiveDirectory module
    Defaults to current user's domain
    Allows for selection of a named domain controller, or the first available
    Uses current user unless -Credential is specified
    Returns a PSObject or Boolean
            
.NOTES
    Name    : Function_Test-PKADConnectio.ps1
    Author  : Paula Kingsley
    Created : 2017-07-24
    Version : 01.00.0000
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2019-06-27 - Created script

.EXAMPLE
    PS C:\> Test-PKADConnection -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value              
        ---              -----              
        Verbose          True               
        ADDomain         corp.net  
        Server                              
        Credential                          
        Boolean          False              
        Quiet            False              
        PipelineInput    False              
        ParameterSetName __AllParameterSets 
        ComputerName     jbloggs-05122    
        ScriptName       Test-PKADConnection
        ScriptVersion    1.0.0              

        BEGIN: Test Active Directory domain connectivity

        Get Active Directory domain
        Connected to Active Directory domain 'corp.net' as 'gmacluskie'
        Get first available Active Directory domain controller
        Got first available Domain Controller

        Domain    : corp.net
        DomainDN  : DC=corp,DC=net
        Server    : OAKWINDOMP001.corp.net
        UserName  : gmacluskie
        SiteName  : US-London
        IsSuccess : True
        Messages  : Connected to Active Directory domain 'corp.net' as 'gmacluskie'
                    Got first available Domain Controller

        END  : Test Active Directory domain connectivity

.EXAMPLE
    Test-PKADConnection -ADDomain foo.bar -Quiet -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value              
        ---              -----              
        ADDomain         foo.bar            
        Quiet            True               
        Verbose          True               
        Server                              
        Credential                          
        Boolean          False              
        PipelineInput    False              
        ParameterSetName __AllParameterSets 
        ComputerName     PKINGSLEY-05122    
        ScriptName       Test-PKADConnection
        ScriptVersion    1.0.0              


        VERBOSE: BEGIN: Test Active Directory domain connectivity
        VERBOSE: Get Active Directory domain
        VERBOSE: Failed to connect to Active Directory domain 'foo.bar' as 'mmax' (Exception calling "GetDomain" with "1" argument(s): "The specified domain does not exist or cannot be contacted.")


        Domain    : foo.bar
        DomainDN  : Error
        Server    : Error
        UserName  : mmax
        SiteName  : Error
        IsSuccess : False
        Messages  : Failed to connect to Active Directory domain 'foo.bar' as 'mmax' (Exception calling "GetDomain" with "1" argument(s): "The specified domain does not exist or cannot be contacted.")

        VERBOSE: END  : Test Active Directory domain connectivity

.EXAMPLE
    PS C:\> Test-PKADConnection -Server foo.domain.com

        BEGIN: Test Active Directory domain connectivity

        Get Active Directory domain
        Connected to Active Directory domain 'corp.net' as 'gmacluskie'
        Get named Active Directory domain controller
        Failed to get named Domain Controller (Exception calling "GetDomainController" with "1" argument(s): "Domain controller "foo.domain.com" does not exist or cannot be contacted.")

        Domain    : corp.net
        DomainDN  : DC=corp,DC=net
        Server    : foo.domain.com
        UserName  : gmacluskie
        SiteName  : Error
        IsSuccess : False
        Messages  : Connected to Active Directory domain 'corp.net' as 'gmacluskie'
                    Failed to get named Domain Controller (Exception calling "GetDomainController" with "1" argument(s): "Domain controller "foo.domain.com" does not exist or cannot be contacted.")

.EXAMPLE
    PS C:\> Test-PKADConnection -ADDomain trusted.domain.local -Credential $Credential 

        BEGIN: Test Active Directory domain connectivity

        Get Active Directory domain
        Connected to Active Directory domain 'trusted.domain.local' as 'TRUSTED\jbloggs'
        Get first available Active Directory domain controller
        Got first available Domain Controller  'DC14.trusted.domain.local'


        Domain    : trusted.domain.local
        DomainDN  : DC=trusted,DC=domain,DC=local
        Server    : DC14.trusted.domain.local
        UserName  : TRUSTED\jbloggs
        SiteName  : London
        IsSuccess : True
        Messages  : Connected to Active Directory domain 'trusted.domain.local' as 'TRUSTED\jbloggs'
                    Got first available Domain Controller 'DC14.trusted.domain.local'

        END  : Test Active Directory domain connectivity

.EXAMPLE
    PS C:\> Test-PKADConnection -ADDomain trusted.domain.local -Quiet

        Domain    : trusted.domain.local
        DomainDN  : DC=trusted,DC=domain,DC=local
        Server    : DC14.trusted.domain.local
        UserName  : gmacluskie
        SiteName  : London
        IsSuccess : True
        Messages  : Connected to Active Directory domain 'trusted.domain.local' as 'gmacluskie'
                    Got first available Domain Controller

.EXAMPLE
    PS C:\> Test-PKADConnection -Server foreigndc.trust.lan

        BEGIN: Test Active Directory domain connectivity

        Get Active Directory domain
        Connected to Active Directory domain 'corp.net' as 'mmax'
        Get named Active Directory domain controller
        WARNING: Domain Controller 'foreigndc.trust.lan' is in domain 'trust.lan,' not 'corp.net'

        Domain    : corp.net
        DomainDN  : DC=corp,DC=net
        Server    : foreigndc.trust.lan
        UserName  : mmax
        SiteName  : Portland
        IsSuccess : False
        Messages  : Connected to Active Directory domain 'corp.net' as 'mmax'
                    Domain Controller 'foreigndc.trust.lan' is in domain 'trust.lan,' not 'corp.net'

        END  : Test Active Directory domain connectivity

.EXAMPLE
    PS C:\> Test-PKADConnection -ADDomain trusted.domain.local -Quiet -Boolean
    True

#>
[Cmdletbinding()]
Param(
    [Parameter(
        HelpMessage = "Active Directory domain FQDN (default is current user)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$ADDomain,

    [Parameter(
        HelpMessage = "Domain controller name or FQDN (default is first available)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [Parameter(
        HelpMessage = "Valid credentials in domain"
    )]
    [ValidateNotNullOrEmpty()]
    [PSCredential]$Credential,

    [Parameter(
        HelpMessage = "Return True/False only"
    )]
    [switch]$Boolean,

    [Parameter(
        HelpMessage = "Suppress all non-verbose console output"
    )]
    [switch]$Quiet

)

Begin {

    # Version from comment block
    [version]$Version = "01.00.0000"

    # Preference
    $ErrorActionPreference = "Stop"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # If there's no domain but there is a server, default to current user's domain (not setting as parameter default to ensure clean testing)
    If ($CurrentParams.Server -and (-not $CurrentParams.ADDomain)) {
        $CurrentParams.ADDomain = $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
    }

    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ComputerName",$Env:ComputerName)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Functions

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

    # Function to write an error or a verbose message
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)#,[switch]$Quiet = $Quiet)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("$Message")}
        Else {Write-Verbose "$Message"}
    }

    #endregion Functions

    #region Output object

    # Normalize username
    If ($CurrentParams.Credential) {$User = $Credential.UserName}
    Else {$User = $Env:UserName}
    
    $InitialValue = "Error"
    $Output = [PSCustomObject]@{
        Domain    = $ADDomain
        DomainDN  = $InitialValue
        Server    = $InitialValue
        UserName  = $User
        SiteName  = $InitialValue
        IsSuccess = $InitialValue
        Messages  = $InitialValue
    }
    $Messages = @()
    
    #endregion Output object

    #region Splats

    # Splat for Write-Progress
    $Activity = "Test Active Directory domain connectivity"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        PercentComplete  = $Null
    }

    #endregion Splats

    $Msg = "BEGIN: $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}
Process {
    
    $Total = 2
    $Current = 0

    # Get the domain/DN
    
    Try {
        $Msg = "Get Active Directory domain"
        $Msg | Write-MessageInfo -FGColor White

        $Current ++
        $Param_WP.CurrentOperation = $Msg
        $Param_WP.PercentComplete = ($Current/$Total*100)
        Write-Progress @Param_WP

        If ($Credential) {
            $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$ADDomain,$Credential.UserName,$Credential.GetNetworkCredential().Password) -ErrorAction Stop   
        }
        Else {
            $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$ADDomain) -ErrorAction Stop
        }
        
        If ($DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)) {
            
            $Output.DomainDN = $DomainObj.GetDirectoryEntry() | Select -ExpandProperty DistinguishedName
            $Msg = "Connected to Active Directory domain '$($DomainObj.Name)' as '$User'"
            $Msg | Write-MessageInfo -FGColor Green
            $Messages += $Msg

            Try {
                
                If ($CurrentParams.Server) {
                    
                    $Msg = "Get named Active Directory domain controller"    
                    $Msg | Write-MessageInfo -FGColor White

                    $Current ++
                    $Param_WP.CurrentOperation = $Msg
                    $Param_WP.PercentComplete = ($Current/$Total*100)
                    Write-Progress @Param_WP

                    $Output.Server = $Server

                    If ($Credential) {
                        $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer",$Server,$Credential.UserName,$Credential.GetNetworkCredential().Password) -ErrorAction Stop
                    }
                    Else {
                        $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("DirectoryServer",$Server) -ErrorAction Stop
                    }

                    If ($DC = [System.DirectoryServices.ActiveDirectory.DomainController]::GetDomainController($Context)) {
                        $Output.Server = $DC.Name
                        $Output.SiteName = $DC.SiteName
                        If ($DC.Domain.Name -eq $ADDomain) {
                            $Output.IsSuccess = $True
                            $Msg = "Verified Domain Controller '$($Server)' in domain '$($DomainObj.Name)'"
                            $Msg | Write-MessageInfo -FGColor Green
                            $Messages += $Msg
                        }
                        Else {
                            $Output.IsSuccess = $False
                            $Msg = "Domain Controller '$($DC.Name)' is in domain '$($DC.Domain),' not '$($DomainObj.Name)'"
                            Write-Warning $Msg
                            $Messages += $Msg
                        }
                    }
                    Else {
                        $Output.IsSuccess = $False
                        $Msg = "Failed to verify Domain Controller '$($Server)' in domain '$($DomainObj.Name)'"
                        $Msg | Write-MessageError
                        $Messages += $Msg
                    }
                }
                Else {
                    
                    $Msg = "Get first available Active Directory domain controller"    
                    $Msg | Write-MessageInfo -FGColor White

                    $Current ++
                    $Param_WP.CurrentOperation = $Msg
                    $Param_WP.PercentComplete = ($Current/$Total*100)
                    Write-Progress @Param_WP

                    If ($DC = [System.DirectoryServices.ActiveDirectory.DomainController]::findone($Context)) {
                        $Output.Server = $DC.Name
                        $Output.SiteName = $DC.SiteName
                        $Output.IsSuccess = $True
                        $Msg = "Got first available Domain Controller '$($DC.Name)'"
                        $Msg | Write-MessageInfo -FGColor Green
                        $Messages += $Msg
                    }
                    Else {
                        $Output.IsSuccess = $False
                        $Msg = "Failed to get first available Domain Controller"
                        $Msg | Write-MessageError
                        $Messages += $Msg
                    }
                }
            }
            Catch {
                $Output.IsSuccess = $False
                If ($Server) {$Msg = "Failed to get named Domain Controller ($Server)"}
                Else {$Msg = "Failed to get Domain Controller"}
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                $Msg | Write-MessageError
                $Messages += $Msg
            }
        }
        Else {
            $Output.IsSuccess = $False
            $Msg = "Failed to connect to Active Directory domain '$($ADDomain)' as '$User'"
            $Msg | Write-MessageError
            $Messages += $Msg
        }
    }
    Catch {
        $Output.IsSuccess = $False
        $Msg = "Failed to connect to Active Directory domain '$($ADDomain)' as '$User'"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
        $Msg | Write-MessageError
        $Messages += $Msg
    }

    If ($Boolean.IsPresent) {
        $Output.IsSuccess
    }
    Else {
        $Output.Messages = $Messages -join("`n")
        Write-Output $Output
    }
}
End {
    
    $Msg = "END  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title
    Write-Progress -Activity $Activity -Completed

}
} #end Test-PKADConnection