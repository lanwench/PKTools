#requires -Version 4
Function New-PKFakeIdentity {
<#
.SYNOPSIS
    Generates one or more random fake identities via the randomuser.me API

.DESCRIPTION
    Calls the randomuser.me API via Invoke-WebRequest to generate fake identities
    Returns a rich set of identity fields, or just basic details if -Basic is specified
    Number of identities is configurable via -Count

.OUTPUTS
    PSObject

.NOTES
    Name    : Function_New-PKFakeIdentity.ps1
    Created : 2024-04-03
    Author  : Paula Kingsley
    Version : 01.01
    History :

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00 - 2024-04-03 - Created script
        v01.01 - 2026-06-17 - Cosmetic updates; converted version format to major.minor

.LINK
    https://publicapis.io/random-user-api


.EXAMPLE
    PS C:\> New-PKFakeIdentity -Verbose
        VERBOSE: PSBoundParameters: 

        Key           Value
        ---           -----
        Verbose       True
        Count         1
        Basic         False
        ScriptName    New-PKFakeIdentity
        ScriptVersion 01.01

        VERBOSE: Running Invoke-WebRequest to URI https://randomuser.me/api/?results=1 to generate 1 fake identities
                                                                                                                                
        Title     : Mrs
        FullName  : Melania Perišić
        GivenName : Melania
        Surname   : Perišić
        Username  : organiclion209
        Email     : melania.perisic@example.com
        Phone     : 037-4438-143
        Gender    : female
        Photo     : https://randomuser.me/api/portraits/thumb/women/64.jpg
        DOB       : 1989-10-19 05:35:20Z
        Age       : 34
        City      : Surdulica
        Country   : Serbia


.EXAMPLE    
    PS C:\> New-PKFakeIdentity -Count 5 -Basic
                                                                                                                        
        GivenName Surname    Username          Email,
        --------- -------    --------          ------
        Boleslav  Burbelo    organicostrich504 boleslav.burbelo@example.com
        Danko     Terzić     organicfish813    danko.terzic@example.com
        Kasper    Pedersen   smallkoala838     kasper.pedersen@example.com
        Elizabeth Graves     purplesnake341    elizabeth.graves@example.com
        Elia      Carpentier organickoala898   elia.carpentier@example.com


#>
Param (
    [Parameter(
        HelpMessage = "Number of fake identities to generate (default is 1)"
    )]
    [ValidateNotNullOrEmpty()]
    [int]$Count = 1,

    [Parameter(
        HelpMessage = "Return name, email, and usernameonly"
    )]
    [switch]$Basic
)
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01"

    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } |
    Where-Object { Test-Path variable:$_ } | ForEach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $URI = "https://randomuser.me/api/?results=$Count"
    $Msg = "Running Invoke-WebRequest to URI $URI to generate $Count fake identities"
    Write-Verbose "[BEGIN: $ScriptName] $Msg"
}
Process {
    
    Try {
        $People = ((Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction SilentlyContinue -Verbose:$False).Content | ConvertFrom-JSON).Results 
        If ($Basic.IsPresent) {
            $People | Select-Object @{N="GivenName";E={$_.name.First}},
            @{N="Surname";E={$_.name.Last}},
            @{N="Username";E={$_.login.username}},
            @{N="Email";E={$_.Email}}
        }
        Else {
            $People | Select-Object @{N="Title";E={$_.Name.Title}},
            @{N="FullName";E={"$($_.name.First) $($_.name.Last)"}},
            @{N="GivenName";E={$_.name.First}},
            @{N="Surname";E={$_.name.Last}},
            @{N="Username";E={$_.login.username}},
            @{N="Email";E={$_.Email}},
            @{N="Phone";E={$_.Phone}},
            @{N="Gender";E={$_.Gender}},
            @{N="Photo";E={$_.picture.thumbnail}},
            @{N="DOB";E={Get-Date $_.dob.date -f u}},
            @{N="Age";E={$_.dob.age}},
            @{N="City";E={$_.location.City}},
            @{N="Country";E={$_.location.Country}}
        }
    }
    Catch {
        Throw $_.Exception.Message
    }
}
End {
    Write-Verbose "[END: $ScriptName] Script ran successfully"
}
} #end

