#requires -Version 4
Function Get-PKDadJoke {
<#
.SYNOPSIS
    Fetches random dad jokes from icanhazdadjoke.com

.DESCRIPTION
    Uses Invoke-WebRequest to call the icanhazdadjoke.com API
    Formats multi-part jokes as Q:/A: when a question mark is detected
    Returns a PSObject with ID and Joke properties
    Supports -WhatIf and -Confirm because this is a Very Serious Operation

.NOTES
    Name    : Function_Get-PKDadJoke.ps1
    Created : 2024-04-03
    Author  : Paula Kingsley
    Version : 02.01
    History :

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00 - 2024-04-03 - Created script
        v02.00 - 2025-12-29 - Fixes issue with URI, simplified output/removed inner function
        v02.01 - 2026-06-17 - Cosmetic updates; corrected synopsis/description; converted version format to major.minor

.OUTPUTS
    PSCustomObject

.PARAMETER URI
    URI for the API call (default is https://icanhazdadjoke.com)

.PARAMETER Count
    Number of jokes to return (1-10; default is 1)

.EXAMPLE
    PS C:\> Get-PKDadJoke -Verbose
    Retrieves a single random dad joke from icanhazdadjoke.com.

.EXAMPLE
    PS C:\> Get-PKDadJoke -Count 3
    Retrieves up to 3 random dad jokes.

#>

    [CmdletBinding(SupportsShouldProcess,ConfirmImpact = "High")]
    Param(
        [Parameter(Position = 0,HelpMessage = "URI to use for API call")]
        [string]$URI = "https://icanhazdadjoke.com",

        [Parameter(HelpMessage = "Number of jokes to return")]
        [ValidateRange(1,10)]
        [int]$Count = 1
    )
    Begin {
        # Current version (please keep up to date from comment block)
        [version]$Version = "02.01"

        # How did we get here?
        $ScriptName = $MyInvocation.MyCommand.Name
        $Source = $PSCmdlet.ParameterSetName
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } |
        Where-Object { Test-Path Variable:$_ } | ForEach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ParameterSetName", $Source)
        $CurrentParams.Add("PipelineInput", $PipelineInput)
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
        
        function Format-DadJoke {
            param([Parameter(ValueFromPipeline)][string]$Joke)
            Begin {}
            Process {
            if ($Joke -match '\?') {
                $parts = $Joke.Split('?',2)
                $Q = ($parts[0] + '?').Trim()
                $A = $parts[1].Trim()
                Write-Output "Q: $Q`nA: $A"
            }
            Else {Write-Output $Joke}
            }
        }


        $Msg = "Use Invoke-WebRequest and to get some really stupid jokes!"
        Write-Verbose "[BEGIN: $ScriptName] $Msg"
    }
    Process {  

        [switch]$Continue = $True
        If (-not $PSCmdlet.ShouldProcess($Env:ComputerName,"Are you sure you want to proceed? Have you really thought about this?")) {
            Write-Output "Cancelling because for once an end user has demonstrated good common sense."
            [switch]$Continue = $False
        }
        If ($Continue.IsPresent) {
            Try {
                
                $Msg = "Getting $Count joke(s) from $URI"
                Write-Verbose $Msg
                [object[]]$Response = Invoke-WebRequest -Headers @{Accept="application/json"} -UserAgent "PowerShell script to show API examples" -Uri $URI -UseBasicParsing -Verbose:$false 
                $Results = $Response.Content | ConvertFrom-Json 
                $Results | Select-Object @{N="ID";E={$_.id}},@{N="Joke";E={$_.joke}}
            }
            Catch {
                Throw $_.Exception.Message
            }
        }
        Else {return}
    }
    End {
        Write-Verbose "[END: $ScriptName] Script ran successfully"
    }
} # end Get-PKDadJoke



Function Get-PKDadJoke2 {
<#
.SYNOPSIS
    Retrieves dad jokes from the icanhazdadjoke API.

.DESCRIPTION
    Supports random jokes or searching. Allows returning multiple jokes
    (up to 30 for search, up to 10 for random).

.PARAMETER Count
    Number of jokes to return (1–10 for random, 1–30 for search).

.PARAMETER Search
    A search term to filter jokes. Uses the /search endpoint.

.EXAMPLE
    Get-PKDadJoke -Count 5
    Returns 5 random jokes.

.EXAMPLE
    Get-PKDadJoke -Search "dog" -Count 3
    Returns 3 jokes containing the word "dog".
#>

    [CmdletBinding(SupportsShouldProcess,ConfirmImpact = "High")]
    Param(
        [Parameter(Position = 0)]
        [string]$URI = "https://icanhazdadjoke.com",

        [Parameter()]
        [ValidateRange(1,30)]
        [int]$Count = 5,

        [Parameter()]
        [string]$Search = "fish"
    )

    Begin {
        [version]$Version = "02.00.0000"
        $ScriptName = $MyInvocation.MyCommand.Name
        Write-Verbose "[BEGIN: $ScriptName] Get dad jokes from icanhazdadjoke.com"

        function _FormatDadJoke {
            param([Parameter(ValueFromPipeline)][string]$Joke)
            Begin {}
            Process {
            if ($Joke -match '\?') {
                $parts = $Joke.Split('?',2)
                $Q = ($parts[0] + '?').Trim()
                $A = $parts[1].Trim()
                #Write-Output "Q: $Q`nA: $A"
                [pscustomobject]@{
                    Question = $Q
                    Answer   = $A
                }
            }
            Else {Write-Output $Joke}
            }
        }
    }

    Process {
        if (-not $PSCmdlet.ShouldProcess($Env:ComputerName,"Fetching dad jokes")) {
            Write-Output "You've shown some uncommonly good sense."
            return
        }

        try {
            if ($Search) {
                $SearchUri = "$URI/search?term=$($Search)&limit=$Count"
                Write-Verbose "Searching for '$Search' ($Count joke(s))"
                $Response = Invoke-WebRequest -Headers @{Accept="application/json"} -UserAgent "PowerShell DadJoke Script" -Uri $SearchUri -UseBasicParsing
                $Results = ($Response.Content | ConvertFrom-Json).results
                #Write-Output $Results | Select-Object ID,@{N="Joke";E={$_.Joke | _FormatDadJoke}}
                Write-Output ($Results.Joke | _FormatDadJoke)
            }
            else {
                Write-Verbose "Fetching $Count random joke(s)"
                $Output = @()
                For ($i = 1; $i -le $Count; $i++) {
                    $Response = Invoke-WebRequest -Headers @{Accept="application/json"} -UserAgent "PowerShell DadJoke Script" -Uri $URI -UseBasicParsing
                    $J = $Response.Content | ConvertFrom-Json
                    $Output += [pscustomobject]@{
                        ID   = $J.id
                        Joke = $J.Joke | _FormatDadJoke
                    }
                    Write-Output ($J.Joke | _FormatDadJoke)
                }
                #Write-Output $Output
            }
        }
        Catch {
            Write-Warning $_.Exception.Message
        }
    }
    End {
        Write-Verbose "[END: $ScriptName] Script ran successfully"
    }
}
