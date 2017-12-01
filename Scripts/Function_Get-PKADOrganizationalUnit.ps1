#requires -Version 3
Function Get-PKADOrganizationalUnit {
<#
.SYNOPSIS
    Uses the ADSI type accelerator to return a menu of Organizational Units in an external gridview (no ActiveDirectory module required)

.DESCRIPTION
    Uses the ADSI type accelerator to return a menu of Organizational Units in an external gridview (no ActiveDirectory module required)
    Defaults searchbase to distinguishedname of current computer's domain 
    Allows for a regular expression match for include/exclude
    Accepts pipeline input
    Returns a PSObject
    

.NOTES
    Name    : Function_Get-PKADOrganizationalUnit
    Created : 2017-10-10
    Version : 01.00.0000
    Author  : Paula Kingsley
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-10-10 - Created script

.PARAMETER BaseDN
    Top-level searchbase in domain (distinguishedname format, e.g., 'DC=domain,DC=local')

.PARAMETER MatchPattern
    String to include in match 

.PARAMETER ExcludePattern
    String to exclude in match

.PARAMETER SuppressConsoleOutput
    Don't output non-verbose/non-error output to display

.EXAMPLE
    PS C:\> Get-PKADOrganizationalUnit -Verbose | Format-List

        VERBOSE: PSBoundParameters: 
	
        Key            Value                           
        ---            -----                           
        Verbose        True                            
        SearchBase     DC=domain,DC=local
        MatchPattern                                   
        ExcludePattern                                 
        ScriptName     Get-PKADOrganizationalUnit      
        ScriptVersion  1.0.0                           


        VERBOSE: Verify DC=domain,DC=local
        VERBOSE: Search for Organizational Units
        VERBOSE: Create menu
        VERBOSE: 1 selection(s) made


        CanonicalName     : domain.local/Development/Operations
        DistinguishedName : OU=Operations,OU=Development,DC=domain,DC=local
        Name              : Operations

.EXAMPLE
    PS C:\> Get-PKADOrganizationalUnit -MatchPattern Group -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                           
        ---            -----                           
        MatchPattern   Group                           
        Verbose        True                            
        SearchBase     DC=domain,DC=local
        ExcludePattern                                 
        ScriptName     Get-PKADOrganizationalUnit      
        ScriptVersion  1.0.0                           



        VERBOSE: Verify DC=domain,DC=local
        VERBOSE: Search for Organizational Units
        VERBOSE: Create menu
        VERBOSE: 3 selection(s) made

        CanonicalName                                                               DistinguishedName                                                
        -------------                                                               -----------------                                                
        domain.local/Production/Gracenote/All Groups/Deployment Groups   OU=Deployment Groups,OU=All Groups,OU=Gracenote,OU=Production,...
        domain.local/Production/Gracenote/All Groups/Distribution Groups OU=Distribution Groups,OU=All Groups,OU=Gracenote,OU=Productio...
        domain.local/Production/Gracenote/All Groups/Security Groups     OU=Security Groups,OU=All Groups,OU=Gracenote,OU=Production,DC...




.EXAMPLE
    PS C:\>  Get-PKADOrganizationalUnit -BaseDN "OU=Computers,DC=domain,DC=local" -ExcludePattern Servers -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                           
        ---            -----                           
        Verbose        True                            
        SearchBase     OU=Computers,DC=domain,DC=local
        MatchPattern   Servers
        ExcludePattern OU=Test                       
        ScriptName     Get-PKADOrganizationalUnit      
        ScriptVersion  1.0.0                           

        VERBOSE: Verify DC=domain,DC=local
        VERBOSE: Search for Organizational Units
        VERBOSE: Create menu
        VERBOSE: 3 selection(s) made

        CanonicalName                                  DistinguishedName                                             Name                                                                             
        -------------                                  -----------------                                             -----                                                                                                                               
        domain.local/Production/Computers/VMs/SQL      OU=SQL,OU=VMs,OU=Computers,OU=Production,DC=domain,DC=local   SQL    
        domain.local/Production/Computers/Laptops      OU=Laptops,OU=Computers,OU=Production,DC=domain,DC=local      Laptops 
        domain.local/Development/Migration/Computers   OU=Computers,OU=Migration,OU=Development,DC=domain,DC=local   Computers

#>
[CmdletBinding()]
 Param(
     [parameter(
         Mandatory = $False,
         ValueFromPipeline = $True,
         ValueFromPipelineByPropertyName = $True,
         HelpMessage = "Top-level DistinguishedName for beginning of search"
     )]
     [Alias("Name","DistinguishedName")]
     [ValidateNotNullOrEmpty()]
     [string]$SearchBase = ([adsi]'').distinguishedName,

     [parameter(
         Mandatory=$False,
         HelpMessage = "Pattern to match (e.g., 'OU=foo')"
     )]
     [ValidateNotNullOrEmpty()]
     [string]$MatchPattern,

     [parameter(
         Mandatory=$False,
         HelpMessage = "Pattern to exclude (e.g., 'OU=bar')"
     )]
     [ValidateNotNullOrEmpty()]
     [string]$ExcludePattern,

     [parameter(
         Mandatory=$False,
         HelpMessage = "Suppress any non-verbose/non-error console output"
     )]
     [ValidateNotNullOrEmpty()]
     [switch]$SuppressConsoleOutput
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

    If (($SearchBase -notmatch "DC=") -or ($SearchBase -match "CN=")) {
        $Msg = "Invalid searchbase '$SearchBase'`nPlease enter as 'DC=foo,DC=com' or 'OU=bar,DC=foo,DC=com'"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }

    If ($CurrentParams.MatchPattern) {
        $Include = [Regex]::Escape($MatchPattern)
    }
    If ($CurrentParams.ExcludePattern) {
        $Exclude = [Regex]::Escape($ExcludePattern)
    }

    $ErrorActionPreference = "SilentlyContinue"

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}
}

Process {

    $Msg = "Verify $SearchBase"
    Write-Verbose $Msg
    $Activity = $Msg
    Write-Progress -Activity $Activity

    # Need to speed this up. It's painfully slow if invalid.
    If (-not ($Null = [adsi]::Exists("LDAP://$SearchBase"))) {
        $Msg = "Invalid entry $SearchBase"
        $Host.UI.WriteErrorLine($Msg)
    }
    Else {
        Try {
            
            $Msg = "Search for Organizational Units"
            Write-Verbose $Msg
            $Activity = $Msg
            Write-Progress -Activity $Activity

            $ObjDomain = [adsi]"LDAP://$SearchBase"
            $Searcher = New-Object DirectoryServices.DirectorySearcher -EA Stop
            $FilterStr = "(objectCategory=organizationalUnit)"
            $Searcher.Filter = $FilterStr
            $Searcher.SearchRoot = $objDomain
            $Searcher.PageSize = 1000
            $Null = $Searcher.PropertiesToLoad.Add("name")
            $Null = $Searcher.PropertiesToLoad.Add("distinguishedname")
            $Null = $Searcher.PropertiesToLoad.Add("canonicalname")
            $AllOUs = $Searcher.FindAll() 
            
            $Results = $AllOUs
            If ($CurrentParams.MatchPattern) {$Results = $Results | Where-Object {$_.Path -match $Include } }
            If ($CurrentParams.ExcludePattern) {$Results = $Results | Where-Object {$_.Path -notmatch $Exclude } }
            
            If ($Results.Count -gt 0) {
                $Msg = "Create menu"
                Write-Verbose $Msg
                $Activity = $Msg
                Write-Progress -Activity $Activity

                [array]$Selection = ($Results | 
                    Select-Object -Expand Properties | 
                        Select @{n='CanonicalName';e={$_.canonicalname}},@{N="DistinguishedName";E={$_.distinguishedname}},@{n='Name';e={$_.name}} | 
                            Sort CanonicalName | 
                                Out-Gridview -OutputMode Multiple -Title "Please select one or more organizational units")
            
                If ($Selection.Count -gt 0) {
                    $Msg = "$($Selection.Count) selection(s) made"
                    Write-Verbose $Msg
                    Write-Output $Selection
                }
                Else {
                    $Msg = "No selection made"
                    Write-Output $Msg
                }
            }
            Else {
                $Msg = "No organizational units found matching your selection"
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        Catch {
            $_.Exception.Message
        }
    }
    Else {
        $Msg = "Invalid entry $SearchBase"
        $Host.UI.WriteErrorLine($Msg)
    }
}
End {
    Write-Progress -Activity $Activity -Completed
}
} #end Get-PKADOrganizationalUnit
