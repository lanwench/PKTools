#requires -version 4
Function Get-PKSitesAndSubnets {
<#
.SYNOPSIS
    Returns AD sites and subnets using .NET

.DESCRIPTION
    Returns AD sites and subnets using .NET
    Returns PSObject or outputs to console with pretty colors
    Uses current user's forest/domain 

.NOTES        
    Name    : Function_Get-PKSitesAndSubnets.ps1
    Created : 2022-03-22
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2022-03-22 - Created script
.EXAMPLE
    PS C:\> Get-PKSitesAndSubnets -OutputTo PSObject -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                
        ---           -----                
        OutputTo      PSObject             
        Verbose       True                 
        PipelineInput False                
        ScriptName    Get-PKSitesAndSubnets
        ScriptVersion 1.0.0                

        VERBOSE: Getting current user's Active Directory forest and sites
        VERBOSE: Found 8 site(s) in forest domain.local


        Forest   : domain.local
        Domains  : {domain.local}
        Site     : LATAM
        Location : Rio de Janeiro, Brazil
        Subnet   : {10.62.0.0/16}
        Servers  : {latamdc01.domain.local, latamdc02.domain.local}

        Forest   : domain.local
        Domains  : {domain.local}
        Site     : NAMW
        Location : Olympia, WA
        Subnet   : {10.61.0.0/16}
        Servers  : {namwdc01.domain.local, namwdc04.domain.local}

        Forest   : domain.local
        Domains  : {domain.local}
        Site     : NAME
        Location : Richmond, VA
        Subnet   : {10.207.216.0/22, 10.6.2.0/23}
        Servers  : {namedc01.domain.local, namedc04.domain.local, namedc05.domain.local}

        Forest   : domain.local
        Domains  : {domain.local}
        Site     : APAC
        Location : Sydney, Aus
        Subnet   : {10.150.128.0/18}
        Servers  : {apacdc01.domain.local}


.EXAMPLE
    PS C:\> Get-PKSitesAndSubnets -OutputTo ConsoleOnly

        ---------------------------------------------------
          Sites/Subnets/Servers in forest domain.local
        ---------------------------------------------------

        [site] NAME
	        [subnet] 10.207.216.0/22
	        [subnet] 10.234.88.128/25
	        [server] namedc01.domain.local
	        [server] namedc04.domain.local
            [server] namedc05.domain.local

        [site] NAMW
	        [subnet] 10.61.0.0/16
	        [server] namwdc01.domain.local
	        [server] namwdc04.domain.local

        [site] LATAM
	        [subnet] 10.62.0.0/16
	        [server] latamdc01.domain.local
	        [server] latamdc02.domain.local

        [site] APAC
	        [subnet] 10.150.128.0/18
	        [server] apacdc01.domain.local
    

#>
[CmdletBinding(DefaultParameterSetName = "PSObject")]
Param(
    
    [Parameter(
        ParameterSetName = "PSObject",
        HelpMessage = "Output sites, subnets, and servers to console with pretty colors"
    )]
    [switch]$PSObject,

    [Parameter(
        ParameterSetName = "ConsoleOnly",
        HelpMessage = "Output sites, subnets, and servers to console with pretty colors"
    )]
    [switch]$ConsoleOnly,

    [Parameter(
        ParameterSetName = "PSObject",
        HelpMessage = "Output properties with collections to arrays with Output sites, subnets, and servers to console with pretty colors"
    )]
    [switch]$CollectionsToStrings,

    [Parameter(
        ParameterSetName = "PSObject",
        HelpMessage = "If -CollectionsToStrings is selected, join properties using a specific delimiter (default is a comma & space)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Comma","Semicolon","LineBreak")]
    [string]$CollectionDelimiter = "Comma"
)
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #$Delim = ', '# '`n'
    $Delim = Switch ($CollectionDelimiter) {
        Comma     {", "}
        Semicolon {"; "}
        LineBreak {"`n"}
    }

    # Console output
    $Activity = "Get Active Directory sites & subnets"
    $Msg = "[BEGIN: $Scriptname] $Activity" 
    Write-Verbose $Msg
}
Process {
    
    Write-Verbose "Getting current user's Active Directory forest and sites"
    If ($Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()) { 
        
        $Sites = $Forest.Sites
        Write-Verbose "Found $($Sites.Count) site(s) in forest $($Forest.Name)"

        Switch ($Source) {
            ConsoleOnly  {
                Write-host "---------------------------------------------------"
                Write-Host "  Sites/Subnets/Servers in forest $($Forest.Name)"
                Write-host "---------------------------------------------------"

                Foreach ($site in ($sites | Sort)) {   
                    write-host "`n[site] $site" -ForegroundColor Cyan  
                    foreach ($subnet in ($site.Subnets | Sort)) {
                        write-host "`t[subnet] $subnet" -ForegroundColor Yellow
                    }
                    foreach ($server in ($site.Servers | Sort)) {
                        write-host "`t[server] $server"
                    }
                }
            }
            PSObject {
                
                Foreach ($Site in $Sites) {
                    
                    $Domains = $Forest.Domains
                    $Subnets = $Site.Subnets
                    $Servers = $Site.Servers
                    If ($CollectionsToStrings.IsPresent) {
                        $Domains = $Forest.Domains -join($Delim)
                        $Subnets = $Site.Subnets -join($Delim)
                        $Servers = $Site.Servers -join($Delim)
                    }
                    
                    [PSCustomObject]@{
                        Forest   = $Forest.Name
                        Domains  = $Domains
                        Site     = $Site.Name
                        Location = $Site.Location
                        Subnet   = $Subnets
                        Servers  = $Servers
                    } 
                }
            } 
        }

        <#
        Switch ($OutputTo) {
            ConsoleOnly {
            
                Write-host "---------------------------------------------------"
                Write-Host "  Sites/Subnets/Servers in forest $($Forest.Name)"
                Write-host "---------------------------------------------------"

                Foreach ($site in ($sites | Sort)) {   
                    write-host "`n[site] $site" -ForegroundColor Cyan  
                    foreach ($subnet in ($site.Subnets | Sort)) {
                        write-host "`t[subnet] $subnet" -ForegroundColor Yellow
                    }
                    foreach ($server in ($site.Servers | Sort)) {
                        write-host "`t[server] $server"
                    }
                }
            }
            PSObject {
                Foreach ($Site in $Sites) {
                    [PSCustomObject]@{
                        Forest   = $Forest.Name
                        Domains  = $Forest.Domains
                        Site     = $Site.Name
                        Location = $Site.Location
                        Subnet   = $site.subnets # -join("`n")
                        Servers  = $Site.Servers # -join("`n")
                    } 
                }
            }
        }
        #>
    }
    Else {
        Write-Warning "No AD forest/site info available in this session"
    }
}
End {
    
    $Msg = "[END: $Scriptname] $Activity" 
    Write-Verbose $Msg

}
} #end Get-PKSitesAndSubnets