#requires -Version 4
Function Test-PKAccountLockout {
<#
.SYNOPSIS
    A simple little function that uses .NET / ADSI searcher to return the account lockout status of one or more AD users in the current user's domain

.DESCRIPTION
    A simple little function that uses .NET / ADSI searcher to return the account lockout status of one or more AD users in the current user's domain
    Defaults to the current user
    ActiveDirectory module is not required
    Output options are Boolean, PSObject, or simply a message written to the console via Write-Host in pretty colors)
    Accepts pipeline input
    Returns a boolean or nothing (outputs a string to the console)

.NOTES        
    Name    : Function_Test-PKAccountLockout.ps1
    Created : 2022-02-24
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2022-02-24 - Created script    

.EXAMPLE


#>
[CmdletBinding()]
Param (
    [Parameter(
        Position = 0,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$SAMAccountName = $env:USERNAME,

    [Parameter(
        HelpMessage= "Output style: Boolean, Console (via Write-Host), or PSObject (default is Console)"
    )]
    [ValidateSet("Boolean","Console","PSObject")]
    [string]$Output = "Console"
) 
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    If ($Output -eq "Console") {
        $Msg = "When Output is set to Console, additional verbose output will be suppressed, and results will be written to the console only"
        Write-Verbose $Msg
    }
    $adsiSearcher = New-Object DirectoryServices.DirectorySearcher("LDAP://rootdse")
}
Process {
    Foreach ($User in $sAMAccountName) {

        $adsiSearcher.filter = "(&(ObjectCategory=User)(sAMAccountName=$User))"
    
        If ($UserObj = $adsiSearcher.FindOne()) {
    
            If( ([adsi]$UserObj.path).psbase.invokeGet("IsAccountLocked") ){
    
                $Msg = "[$(Get-Date -f "yyyy-MM-dd HH:mm:ss")] Account '$($UserObj.Properties.samaccountname)' is locked out ($(([adsi]$UserObj.path).DistinguishedName))"
                
                Switch ($Output) {
                    Boolean {
                        Write-Warning $Msg
                        $True
                    }
                    Console {
                        Write-Host $Msg -ForegroundColor Red
                    }
                    PSObject {
                        Write-Verbose $Msg
                        [PSCustomObject]@{
                            SAMAccountName    = $($UserObj.Properties.samaccountname)
                            Name              = $($UserObj.Properties.name)
                            IsLockedOut       = $True
                            DistinguishedName = $($UserObj.Properties.distinguishedname)
                        }
                    }
                } #end switch
            }
            Else {
                $Msg = "[$(Get-Date -f "yyyy-MM-dd HH:mm:ss")] Account '$($UserObj.Properties.samaccountname)' is not locked out ($(([adsi]$UserObj.path).DistinguishedName))"

                Switch ($Output) {
                    Boolean {
                        Write-Verbose $Msg
                        $False
                    }
                    Console {
                        Write-Host $Msg -ForegroundColor Green
                    }
                    PSObject {
                        Write-Verbose $Msg
                        [PSCustomObject]@{
                            SAMAccountName    = $($UserObj.Properties.samaccountname)
                            Name              = $($UserObj.Properties.name)
                            IsLockedOut       = $False
                            DistinguishedName = $($UserObj.Properties.distinguishedname)
                        }
                    }
                } # end switch
            }
        } #end if found
        Else {

            $Msg = "[$(Get-Date -f "yyyy-MM-dd HH:mm:ss")] Account '$User' not found"

            Switch ($Output) {
                Boolean {
                    Write-Warning $Msg
                }
                Console {
                    Write-Host $Msg -ForegroundColor Red
                }
                PSObject {
                    Write-Warning $Msg
                    [PSCustomObject]@{
                        SAMAccountName    = $User
                        Name              = "ERROR"
                        IsLockedOut       = "ERROR"
                        DistinguishedName = "ERROR"
                    }
                }  
            } #end switch
        } #end if not found
    } #end foreach
}
} #end Test-PKAccountLockout

  