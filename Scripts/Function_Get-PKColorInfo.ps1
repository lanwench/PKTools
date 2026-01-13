#requires -version 4.0
Function Get-PKColorInfo {

<#
.SYNOPSIS
    Look up color information by Hex code or Name with ANSI color examples, via REST API (default) or local query.

.DESCRIPTION
    Look up color information by Hex code or Name with ANSI color examples, via REST API (default) or local query.
    - Hex Codes: No '#' required. Just type the 3 or 6 digit code. (Asterisks not allowed).
    - Names: 
        - Exact match required by default (e.g., "Blue").
        - Use '*' for wildcard matching (e.g., "Dark*"). 
        - If wildcard returns multiple matches, a selection menu appears unless -DisplayAll is used.
    - Output: Returns a rich object with ANSI color preview.

.NOTES
    Name    : Function_Get-PKColorInfo.ps1
    Author  : Paula Kingsley
    Created : 2026-01-08
    Version : 01.00.0000    
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2026-01-08 - Created script

.PARAMETER InputValue
    One or more color identifiers as Hex codes or color names. Accepts pipeline input.

.PARAMETER DisplayAll
    When specified with wildcard color name matching, displays all matches instead of showing a selection menu.

.PARAMETER LocalLookup
    When specified, performs color lookup using a local color database.

.OUTPUTS
    System.Object
    Returns an object containing color information including Hex code, RGB values, and ANSI color preview.

.EXAMPLE
    PS C:\> Get-PKColorInfo ff0000, 0047AB
    Looks up two colors by their Hex codes and returns their information.

.EXAMPLE
    PS C:\> "Dark*", "Blue" | Get-PKColorInfo -DisplayAll -LocalLookup
    Looks up colors matching "Dark*" pattern and exact match "Blue" using local lookup, displaying all results.

.EXAMPLE
    PS C:\> Get-PKColorInfo -Color dark*green -DisplayAllMatches -LocalLookup -Small

        Swatch  Name           Hex
        ------  ----           ---
        ███████ DarkGreen      006400
        ███████ DarkOliveGreen 556B2F
        ███████ DarkSeaGreen   8FBC8F


.EXAMPLE
    PS C:\> "Dark*", "Blue" | Get-PKColorInfo -DisplayAll -LocalLookup  | Format-Table -AutoSize

        Swatch  Name           Hex    RGB                Source
        ------  ----           ---    ---                ------
        ███████ DarkBlue       00008B rgb(0, 0, 139)     System.Drawing.Color
        ███████ DarkCyan       008B8B rgb(0, 139, 139)   System.Drawing.Color
        ███████ DarkGoldenrod  B8860B rgb(184, 134, 11)  System.Drawing.Color
        ███████ DarkGray       A9A9A9 rgb(169, 169, 169) System.Drawing.Color
        ███████ DarkGreen      006400 rgb(0, 100, 0)     System.Drawing.Color
        ███████ DarkKhaki      BDB76B rgb(189, 183, 107) System.Drawing.Color
        ███████ DarkMagenta    8B008B rgb(139, 0, 139)   System.Drawing.Color
        ███████ DarkOliveGreen 556B2F rgb(85, 107, 47)   System.Drawing.Color
        ███████ DarkOrange     FF8C00 rgb(255, 140, 0)   System.Drawing.Color
        ███████ DarkOrchid     9932CC rgb(153, 50, 204)  System.Drawing.Color
        ███████ DarkRed        8B0000 rgb(139, 0, 0)     System.Drawing.Color
        ███████ DarkSalmon     E9967A rgb(233, 150, 122) System.Drawing.Color
        ███████ DarkSeaGreen   8FBC8F rgb(143, 188, 143) System.Drawing.Color
        ███████ DarkSlateBlue  483D8B rgb(72, 61, 139)   System.Drawing.Color
        ███████ DarkSlateGray  2F4F4F rgb(47, 79, 79)    System.Drawing.Color
        ███████ DarkTurquoise  00CED1 rgb(0, 206, 209)   System.Drawing.Color
        ███████ DarkViolet     9400D3 rgb(148, 0, 211)   System.Drawing.Color
        ███████ Blue           0000FF rgb(0, 0, 255)     System.Drawing.Color

#>
    [CmdletBinding(DefaultParameterSetName = "Medium")]
    Param(
        [Parameter(
            Mandatory=$true, 
            Position=0, 
            ValueFromPipeline=$true, 
            ValueFromPipelineByPropertyName=$true, 
            HelpMessage="One or more hex codes (no # needed) or color names (wildcards allowed)"
        )]
        [string[]]$Color,

        [Parameter(
            HelpMessage="Maximum number (1-50) of ambiguous matches to display for selection"
        )]
        [ValidateRange(1,50)]
        [int]$MaxLookups = 10,

        [Parameter(
            ParameterSetName = "Large",
            HelpMessage = "Return all properties"
        )]
        [switch]$Large,

        [Parameter(
            ParameterSetName = "Medium",
            HelpMessage = "Return subset of properties (this is the default output)"
        )]
        [switch]$Medium,
        
        [Parameter(
            ParameterSetName = "Small",
            HelpMessage = "Return minimal properties (name, hex, swatch)"
        )]
        [switch]$Small,
        

        [Parameter(
            HelpMessage="Skip API call and use local System.Drawing"
        )]
        [switch]$LocalLookup,

        [Parameter(
            HelpMessage="Display all matches instead of showing a selection menu"
        )]
        [switch]$DisplayAllMatches
    )

    Begin {
        
        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # Show our settings
        $ScriptName = $MyInvocation.MyCommand.Name
        $CurrentParams = $PSBoundParameters
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("PipelineInput", $PipelineInput)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        If ($Small.IsPresent) { $SelectProps = "Swatch,Name,Hex" -split "," }
        Elseif ($Large.IsPresent) { $SelectProps = "*" }
        Else {$SelectProps = "Swatch,Name,Hex,RGB,ExactMatch" -split "," }

        # Optimization: Pre-define Regex pattern
        $HexPattern = "^([a-fA-F0-9]{3}|[a-fA-F0-9]{6})$"

        # Load System.Drawing for local name lookups/conversions
        Add-Type -AssemblyName System.Drawing
        $Esc = [char]27
        $AllColors = [Enum]::GetNames([System.Drawing.KnownColor])

        If ($LocalLookup.IsPresent) {
            Write-Verbose "[BEGIN: $ScriptName] Invoke API call to ColorAPI to look up colors by hex code or name"
        }
        Else {
            Write-Verbose "[BEGIN: $ScriptName] Look up local system color information by hex code or name"
        }
    }

    Process {
        Foreach ($Item in $Color) {
            # 1. Clean Input
            If ([string]::IsNullOrWhiteSpace($Item)) { Continue }

            # Remove '#' if provided, trim whitespace, determine whether it's a hex code or not
            $CleanQuery = $Item.Trim().TrimStart('#')
            $IsHex = $CleanQuery -match $HexPattern
            
            # List to hold targets for processing: @{ Name="Blue"; Hex="0000FF" }
            $ProcessQueue = @()

            # 2. Determine Targets
            If ($IsHex) {
                Write-Verbose "[$CleanQuery] Input detected as hex code"
                $ProcessQueue += [PSCustomObject]@{ Name = $null; Hex = $CleanQuery }
            }
            Else {
                Write-Verbose "[$CleanQuery] Input detected as text string"
                $FoundColors = @()
                If ($CleanQuery.Contains('*')) {
                    Write-Verbose "[$CleanQuery] Wildcard detected ...performing fuzzy search"
                    # Optimization: Use .Where() method instead of pipeline
                    $FoundColors = @($AllColors.Where({ $_ -like $CleanQuery }) | Sort-Object)
                }
                Else {
                    Write-Verbose "[$CleanQuery] No wildcard detected; performing exact match check"
                    # Optimization: Use .Where() method instead of pipeline
                    $FoundColors = @($AllColors.Where({ $_ -eq $CleanQuery }))
                }

                # Handle Search Results
                If ($FoundColors.Count -eq 0) {
                    Write-Warning "[$CleanQuery] Color not found! (Use '*' for partial matching, e.g., 'Dark*')"
                    Continue
                }
                ElseIf ($FoundColors.Count -gt 1) {
                    If ($DisplayAllMatches) {
                        Write-Verbose "[$CleanQuery] $($FoundColors.Count) matches found using $MaxLookups maximum lookups; adding all to queue"
                        Foreach ($ColorName in $FoundColors) {
                            $kColor = [Enum]::Parse([System.Drawing.KnownColor], $ColorName)
                            $SysColor = [System.Drawing.Color]::FromKnownColor($kColor)
                            $HexVal = "{0:X2}{1:X2}{2:X2}" -f $SysColor.R, $SysColor.G, $SysColor.B
                            $ProcessQueue += [PSCustomObject]@{ Name = $ColorName; Hex = $HexVal }
                        }
                    }
                    Else {
                        $DisplayCount = [Math]::Min($FoundColors.Count, $MaxLookups)
                        $HasMore = $FoundColors.Count -gt $MaxLookups
                        
                        Write-Host "[$CleanQuery] $($FoundColors.Count) matches found...showing top $DisplayCount`:" -ForegroundColor Cyan
                        
                        # Loop through and display the menu
                        For ($i = 0; $i -lt $DisplayCount; $i++) {
                            $NameCandidate = $FoundColors[$i]
                            $c = [System.Drawing.Color]::FromName($NameCandidate)
                            # Create ANSI Swatch
                            $MenuSwatch = "$Esc[38;2;$($c.R);$($c.G);$($c.B)m████$Esc[0m"
                            Write-Host " [$i] $MenuSwatch $NameCandidate"
                        }
                        If ($HasMore) {Write-Host "...and $($FoundColors.Count - $MaxLookups) more ... try a more specific query or change -MaxLookups" -ForegroundColor Gray}

                        # Interactive Prompt
                        $Selection = Read-Host "`nEnter number to select (or Press Enter to skip)"
                        
                        If ($Selection -match "^\d+$" -and [int]$Selection -lt $DisplayCount) {
                            Write-Verbose "[$CleanQuery] Converting names to hex codes"
                            $ColorName = $FoundColors[[int]$Selection]
                            $kColor = [Enum]::Parse([System.Drawing.KnownColor], $ColorName)
                            $SysColor = [System.Drawing.Color]::FromKnownColor($kColor)
                            $HexVal = "{0:X2}{1:X2}{2:X2}" -f $SysColor.R, $SysColor.G, $SysColor.B
                            $ProcessQueue += [PSCustomObject]@{ Name = $ColorName; Hex = $HexVal }
                        }
                        Else {
                            Write-Warning "[$CleanQuery] Selection cancelled for '$Item'"
                            Continue
                        }
                    }
                }
                Else {
                    Write-Verbose "[$CleanQuery] Converting matching local color name to hex code"
                    $ColorName = $FoundColors[0]
                    $kColor = [Enum]::Parse([System.Drawing.KnownColor], $ColorName)
                    $SysColor = [System.Drawing.Color]::FromKnownColor($kColor)
                    $HexVal = "{0:X2}{1:X2}{2:X2}" -f $SysColor.R, $SysColor.G, $SysColor.B
                    $ProcessQueue += [PSCustomObject]@{ Name = $ColorName; Hex = $HexVal }
                }
            }

            # 3. Process Queue (API or Local)
            Foreach ($Target in $ProcessQueue) {
                $TargetHex = $Target.Hex
                $TargetName = $Target.Name
                
                If ($LocalLookup) {
                    Try {
                        Write-Verbose "[$CleanQuery] Performing local lookup for #$TargetHex"
                        $ColorObj = [System.Drawing.ColorTranslator]::FromHtml("#$TargetHex")
                        $R = $ColorObj.R; $G = $ColorObj.G; $B = $ColorObj.B
                        
                        # Create ANSI Swatch
                        $Swatch = "$Esc[38;2;${R};${G};${B}m███████$Esc[0m"
                        
                        $Output = [PSCustomObject]@{
                            Swatch     = $Swatch
                            Input      = $Item  # Optimization: Use specific item
                            Name       = if ($TargetName) { $TargetName } else { $ColorObj.Name }
                            Hex        = $TargetHex
                            RGB        = "rgb($R, $G, $B)"
                            Method     = "Local Lookup"
                            ExactMatch = $true
                            Source     = "System.Drawing.Color"
                        }
                        Write-Output ($Output | Select-Object $SelectProps)
                    }
                    Catch {
                        Write-Error "[$CleanQuery] Local lookup failed for #$TargetHex`: $_"
                    }
                }
                Else {
                    $Url = "https://www.thecolorapi.com/id?hex=$TargetHex"
                    Try {
                        Write-Verbose "[$CleanQuery] Initiating REST API Call to $Url"
                        $Resp = Invoke-RestMethod -Uri $Url -Method Get -Verbose:$False
                        
                        # Extract RGB for ANSI Block (Final Output)
                        $r = $Resp.rgb.r
                        $g = $Resp.rgb.g
                        $b = $Resp.rgb.b
                        $Swatch = "$Esc[38;2;$r;$g;${b}m███████$Esc[0m"

                        # Determine Display Name
                        $DisplayName = If ($TargetName) { $TargetName } Else { $Resp.name.value }
                        $ApiNameRaw = $Resp.name.value

                        # Build Output Object
                        $Output = [PSCustomObject]@{
                            Swatch     = $Swatch
                            Input      = $Item  
                            Name       = $DisplayName
                            APIName    = $ApiNameRaw 
                            Hex        = $Resp.hex.value
                            RGB        = $Resp.rgb.value
                            HSL        = $Resp.hsl.value
                            Method     = "API query" 
                            ExactMatch = $Resp.name.exact_match_name
                            Image      = $Resp.image.bare
                            Source     = $Url
                        }

                        Write-Output ($Output | Select-Object $SelectProps)
                    }
                    Catch {
                        Write-Error "[$CleanQuery] Failed to retrieve data for #$TargetHex`: $_"
                    }
                }
            }
        } 
    }
End {
    Write-Verbose "[END: $ScriptName] Operation complete"
}    
} # end Get-PKColorInfo



Get-PKColorInfo -Color dark*green -DisplayAllMatches -LocalLookup -Large