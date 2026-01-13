#requires -Version 4
Function Get-PKJOTD {
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
        [Parameter(Position = 0,HelpMessage = "categories to search for (if unspecified, defaults to 'Programming' and 'Miscellaneous')")]
        [ValidateSet("Any","Miscellaneous","Programming","Dark","Pun","Spooky","Christmas")]
        [Alias("Tags")]
        [string[]]$Categories,

        [Parameter(HelpMessage = "Number of jokes to return")]
        [ValidateRange(1,10)]
        [int]$Count = 1,
        
        [Parameter(HelpMessage = "Return two-parters, singles/one-liners, or any")]
        [ValidateSet("TwoPart","Single","Any")]
        [Alias("Style")]
        [string]$Type = "TwoPart",

        [Parameter(HelpMessage = "Don't filter out inappropriate content (except racist & sexist material, because, come on, nope)")]
        [switch]$IncludeNFSW
    )
    Begin {
        # Current version (please keep up to date from comment block)
        [version]$Version = "02.00.0000"

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
        
        
        #region Build URI string        
        <#
        If (-not $Categories) {$URI = "https://v2.jokeapi.dev/joke/Programming,Miscellaneous"}
        Else {$URI = "https://v2.jokeapi.dev/joke/$($Categories -join(","))"}
        If ($IncludeNFSW.IsPresent) {$URI += "?blacklistFlags=nsfw,religious,racist,sexist,explicit,political"}
        Else {$URI += "?blacklistFlags=nsfw,religious,explicit,political"}
        $URI += "&amount=$Count"
        #>

        #endregion Build URI string

        #region Build URI string
        

        # Base categories
        If ($Categories) {$URI = "https://v2.jokeapi.dev/joke/$($Categories -join(','))"}
        Else {$URI = "https://v2.jokeapi.dev/joke/Programming,Miscellaneous"}
        
        # Always blocked
        $Blacklist = @("racist","sexist")

        # Detect if Dark category is used
        $UsingDark = $Categories -contains "Dark"

        # NSFW allowed only when explicitly requested
        if (-not $IncludeNFSW.IsPresent) {$Blacklist += "nsfw"}

        # Religious, political, explicit allowed only if NSFW or Dark
        if (-not $IncludeNFSW.IsPresent -and -not $UsingDark) {$Blacklist += @("religious","political","explicit")}

        # Build final URI
        $URI += "?blacklistFlags=$($Blacklist -join(','))"
        $URI += "&amount=$Count"

        #endregion Build URI string


        $Msg = "Use Invoke-WebRequest and jokeapi.dev to get jokes!"
        Write-Verbose "[BEGIN:  $ScriptName] $Msg"       
    }
    Process {  

        [switch]$Continue = $True
        If ($Categories -contains "Dark") {
            $Msg = "You've selected 'Dark' as a category."
            If (-not $IncludeNFSW.IsPresent) {
                $Msg += " Although you didn't double down by specifying -IncludeNFSW, this could still be gross."
            }
            $Msg += "`nThis script will always *attempt* to filter out at least racist and sexist content, regardless, but this is no guarantee. Continue at your peril."
            Write-Warning $Msg
            If (-not $PSCmdlet.ShouldProcess($Env:ComputerName,"Are you sure  you want to proceed?")) {
                Write-Output "Cancelling because user has demonstrated good common sense."
                [switch]$Continue = $False
            }
        }

        If ($Continue.IsPresent) {
            Try {
                Switch ($Type) {
                    "Any" {}
                    "TwoPart" {$URI = $("$URI&type=twopart")}
                    "Single" {$URI = $("$URI&type=single")}
                }
                
                $Msg = "Getting $Count joke(s) from $URI"
                Write-Verbose $Msg
                [object[]]$Response = Invoke-WebRequest -Uri $URI -UseBasicParsing -Verbose:$false 
                $ResponseStream = $Response.RawContentStream
                $StreamReader = New-Object System.IO.StreamReader($ResponseStream, [System.Text.Encoding]::UTF8)
                $Results = $StreamReader.ReadToEnd() | ConvertFrom-Json 
                If ($Results.Amount -gt 1) {$Results = $Results.Jokes}

                $Results | Select-Object @{N="Category";E={$_.category}},
                    @{N="Type";E={$_.type}},
                    @{N="Safe";E={$_.safe}},
                    @{N="ID";E={$_.id}},
                    @{N="Joke";E={
                        If ($_.type -eq 'single') {$_.joke}
                        Elseif ($_.type -eq 'twopart') {"Q: $($_.setup)`nA: $($_.delivery)"}
                    }}
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
} #end function   