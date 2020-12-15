#requires -version 4
Function Test-PKPasswordPolicy {
<#
.SYNOPSIS
    Tests a string against a domain password policy for length and complexity

.DESCRIPTION
    Tests a string against a domain password policy for length and complexity
    Defaults to current user's domain (can provide either domain or server)
    Accepts pipeline input
    Returns a boolean

.NOTES        
    Name    : Function_Test-PKPasswordPolicy.ps1
    Created : 2020-06-24
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-06-24 - Created script based on Allan Rafuse's original (see link)

    Original author: Allan Rafuse

.LINK
    https://www.checkyourlogs.net/active-directory-password-complexity-check-powershell-mvphour/

#>
[Cmdletbinding(
    DefaultParameterSetName = "Identity"
)]

Param (
    [Parameter(
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True
    )]
    [string]$Password,
        
    [Parameter(
        ParameterSetName = "Identity",
        HelpMessage = "Name of domain, if not providing domain controller"
    )]
    [Alias("ADDomain","Domain")]
    [string]$Identity = $Env:USERDOMAIN,

    [Parameter(
        ParameterSetName = "Server",
        HelpMessage = "Name of domain controller, if not providing domain name"

    )]
    [string]$Server
)
    
Begin {
    
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Password = "****"
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
    If (-not $Password) {
        
        $Password = Read-Host -Prompt "Enter password" -AsSecureString
        If (-not $Password.Length -gt 0) {
            $Msg = "You must provide a password to test"
            Throw $Msg
        }
    }

    Try {
        Switch ($Source) {
            Identity {$PasswordPolicy = Get-ADDefaultDomainPasswordPolicy -Identity $Identity -ErrorAction Stop}
            Server {$PasswordPolicy = Get-ADDefaultDomainPasswordPolicy -Server $Server -ErrorAction Stop}
        }
        $Msg = "Successfully got password policy for domain '$($PasswordPolicy.DistinguishedName)'"
        Write-Verbose $Msg
        
        #$Msg =  "Current settings `n`t$($PasswordPolicy | Format-Table -AutoSize | out-string )"
        $Msg =  "Current settings `n`t$($PasswordPolicy | out-string )"
        Write-Verbose $Msg
    }
    Catch {
        $Msg = "Failed to get domain password policy ($($_.Exception.Message))"
        Throw $Msg
    }
}
Process {

    Foreach ($PW in $Password) {
        
        # Set the flag
        [switch]$Continue = $True

        # Make sure it's long enough
        # The if continue is just for consistency; we could set it at the top, but meh
        If ($Continue.IsPresent) {
            
            # Reset flag
            $Continue = $False 

            # Make sure it's long enough
            If ($Password.Length -ge $PasswordPolicy.MinPasswordLength) {
                $Msg = "Password meets minimum length requirement of $($PasswordPolicy.MinPasswordLength)"
                Write-Verbose $Msg
                $Continue = $True
            }
            Else {
                $Msg = "Password does not meet minimum length requirement of $($PasswordPolicy.MinPasswordLength)"
                Write-Warning $Msg
                $False               
            }
        }

        # Test complexity
        If ($Continue.IsPresent -and $PasswordPolicy.ComplexityEnabled) {

            If ( `   
                 ($PW -cmatch "[A-Z\p{Lu}\s]") `
                -and ($PW -cmatch "[a-z\p{Ll}\s]") `
                -and ($PW -match "[\d]") `
                -and ($PW -match "[^\w]")  
            ) { 
                $Msg = "Password appears to match domain complexity rules"
                Write-Verbose $Msg
                $true
            }
            Else {
                $Msg = "Password does not appear to match domain complexity rules"
                Write-Warning $Msg
                $False
            }
        } 
    } #end foreach
}    
} #end function


