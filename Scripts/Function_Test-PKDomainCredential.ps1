#requires -Version 3
Function Test-PKDomainCredential{
<# 
.SYNOPSIS
    Tests authentication to Active Directory

.DESCRIPTION
    Tests authentication to Active Directory
    Accepts a username (and prompts for password), or a PSCredential object
    AD NetBIOS or DNS domain name can be supplied; defaults to current user's domain, and is 
    ignored if username contains the domain name 
    Returns a Boolean

.NOTES        
    Name    : Function_Test-PKDomainCredential.ps1
    Created : 2020-11-06
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-11-06 - Created script
        
.PARAMETER UserName         
    Domain username, in one of the following formats: username, DOMAIN\username, username@domain.local

.PARAMETER Credential             
    Active Directory PSCredential object, with username in one of the following formats: username, DOMAIN\username, username@domain.local

.PARAMETER ADDomain
    NetBIOS or DNSDomain Active Directory domain name  (default is logged in user's; ignored if username is DOMAIN\user or user@domain.com)


.EXAMPLE
    PS C:\> Do-SomethingCool -ComputerName foo

#> 

[CmdletBinding(DefaultParameterSetName = "User")]
Param(
    
    [Parameter(
        ParametersetName = "User",
        Mandatory = $True,
        HelpMessage = "Domain username, in one of the following formats: username, DOMAIN\username, username@domain.local"
    )]
    [string]$UserName,

    [Parameter(
        ParametersetName = "PSCred",
        Mandatory = $True,
        HelpMessage = "Active Directory PSCredential object, with username in one of the following formats: username, DOMAIN\username, username@domain.local"
    )]
    [PSCredential]$Credential,

    [Parameter(
        HelpMessage = "NetBIOS or DNSDomain Active Directory domain name  (default is logged in user's; ignored if username is DOMAIN\user or user@domain.com)"
    )]
    [ValidateNotNullOrEmpty()]
    [String]$ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,

    [Parameter(
        HelpMessage = "Prefer the -ADDomain value (useful if username does not contain AD domain, or UPN does not match AD domain name)"
    )]
    [switch]$PreferADDomain

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    #[switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    #$CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
}
Process {
    
    $User = $Domain = $Null

    # Parse out the user & domain name
    Switch ($Source) {

        User{

            If ($UserName -match "\\" ) {
                $User = (($UserName -Split "\\")[1]).Trim()
                $Domain = (($UserName -Split "\\")[0]).Trim()
            }
            ElseIf ($UserName -match "@" ) {
                $User = (($UserName -Split "@")[0]).Trim()
                $Domain = (($UserName -Split "@")[1]).Trim()
            }
            Else {
                $User = $UserName
            }
        }

        PSCred {
            
            $UserName = $Credential.UserName
            $Password = $Credential.GetNetworkCredential().Password
            
            If ($UserName -match "\\" ) {
                $User = (($UserName -Split "\\")[1]).Trim()
                $Domain = (($UserName -Split "\\")[0]).Trim()
            }
            ElseIf ($UserName -match "@" ) {
                $User = (($UserName -Split "@")[0]).Trim()
                $Domain = (($UserName -Split "@")[1]).Trim()
            }
        }
    }

    # Make sure we have the domain for context
    If ($Domain) {
        If (-not $PreferADDomain.IsPresent) {
            $Msg = "Domain name '$Domain' found in Username; -PreferADDomain not set to TRUE"
            Write-Verbose $Msg    
        }
        Else {
            $Msg = "-PreferADDomain set to TRUE; will override '$Domain' from Username with ADDomain '$ADDomain'"
            Write-Verbose $Msg
        }
    }
    Else {
        $Domain = $ADDomain
    }

    # Make a credential object for the username
    If (-not $Credential.Username) {
        $Credential = Get-Credential $User -Message "Please enter the password for domain '$Domain' user '$User'"
        If ($Credential.UserName) {
            $Password = $Credential.GetNetworkCredential().Password
        }
        Else {
            $Msg = "No credential object created; please re-run script"
            Write-Error $Msg
        }
    }
    
    # Test
    If ($Credential) {
        
        $Msg = "Test authentication for user '$User' in domain '$Domain'"
        Write-Verbose $Msg

        Try {
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            $PrincipalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, $Domain)

            Try {

                If ($PrincipalContext.ValidateCredentials($User,$Password)) {
                    $Msg = "Successfully authenticated '$User' in '$Domain'"
                    Write-Verbose $Msg
                    $True
                }
                Else {
                    $Msg = "Failed to authenticate '$User' in '$Domain'"
                    Write-Error $Msg
                    $False
                }
            }
            Catch {
                $Msg = "Unknown error ($($_.Exception.Message)"
                Write-Error $Msg
            }

            $Principalcontext.Dispose()
        }
        Catch {
            $Msg = "Authentication failed ($($_.Exception.Message)"
            Write-Error $Msg
        }
        
    }

}
End {}
}