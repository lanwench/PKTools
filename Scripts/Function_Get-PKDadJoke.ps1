#requires -Version 4
Function Get-PKDadJoke {
<#
.SYNOPSIS
    Retrieves jokes from the v2.jokeapi.dev API based on specified parameters

.DESCRIPTION
    The Get-PKJOTD function retrieves jokes from the Joke API based on the specified parameters
    It allows you to specify categories, the number of jokes to return, the style of jokes (two-part, one-liner, or any), and whether to include NSFW (Not Safe for Work) content
    Responses are converted to UTF8 and returned as a PSCustomObject
    This was created as a way to have a little fun with a free API endpoint
    NOTE: It will always filter out the categories 'racist' and 'sexist' but this is no guarantee, so use at your own risk!

.NOTES
    Name    : Function_Get-PKJOTD.ps1
    Created : 2024-04-03
    Author  : Paula Kingsley
    Version : 02.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2024-04-03 - Created script
        v02.00.0000 - 2025-12-29 - Fixes issue with URI, simplified output/removed inner function

.OUTPUTS
    PSCustomObject        

.PARAMETER categories
    Specifies the categories to search for. By default, it searches for jokes with categories "Programming" and "Miscellaneous".

.PARAMETER Count
    Specifies the number of jokes to return. The valid range is between 1 and 10. The default value is 1.

.PARAMETER Style
    Specifies the style of jokes to return. Valid values are "TwoPart" (two-part jokes), "single" (one-liners), or "any" (any style). The default value is "any".

.PARAMETER IncludeNFSW
    Specifies whether to include NSFW (Not Safe for Work) content in the results. By default, NSFW content is filtered out. Sexist/racist categories are always filtered out!

.EXAMPLE
    PS C:\> Get-PKJOTD -categories "Programming" -Count 3 -Style "TwoPart"
    Retrieves 3 two-part jokes with the tag "Programming".

.EXAMPLE
    PS C:\> Get-PKJOTD -Count 5
    Retrieves 5 jokes with any style and any categories.

.EXAMPLE
    PS C:\> Get-PKJOTD -Style single -Count 5 -IncludeNFSW
    Retrieves 3 single jokes and any categories.

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
        [version]$Version = "01.00.0000"

        # How did we get here?
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $CurrentParams = $PSBoundParameters
        #If (-not $CurrentParams.Type) {$Type = $CurrentParams.Type = "TwoPart"}
        $ScriptName = $MyInvocation.MyCommand.Name
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path Variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        $CurrentParams.Add("PipelineInput", $PipelineInput)
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
        Write-Verbose "[BEGIN:  $ScriptName] $Msg"       
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
        Write-Verbose "[END:  $ScriptName]"
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
        Write-Verbose "[BEGIN] Get-PKDadJoke v$Version"

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
        Write-Verbose "[END: $ScriptName]"
    }
}
