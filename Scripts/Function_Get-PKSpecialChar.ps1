#requires -Version 4

Function Get-PKSpecialChar {
<#
.SYNOPSIS
    Retrieves Unicode characters or their corresponding code points based on specified parameters.

.DESCRIPTION
    Retrieves Unicode characters or their corresponding code points based on specified parameters.
    Allows users to retrieve Unicode characters or their code points 
    based on a specified range, code, or character input. It supports selecting by named Unicode ranges 
    and includes options to handle non-printable characters.

.NOTES        
    Name    : Function_Get-PKSpecialChar.ps1
    Created : 2025-05-13
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2025-05-13 - Created script

.INPUTS
    - [string[]] (for the -CodeType' parameter)
    - [int[]] (for the '-Code' parameter)
    - [string[]] (for the '-Char' parameter)

.OUTPUTS
    - [PSCustomObject] containing the following properties:
        - 'Set': The Unicode range name (if applicable).
        - 'Code': The Unicode code point in decimal format.
        - 'Char': The corresponding character.
        - 'Syntax': The PowerShell syntax to represent the character.

.LINK
    https://docs.microsoft.com/en-us/powershell/    

.PARAMETER CodeType
    Specifies named sets of common Unicode characters. Acceptable values include:
    - Arrows
    - BasicLatin
    - Currency
    - LatinExtendedA
    - LatinExtendedB
    - LatinSupplement
    - Mathematical
    - Punctuation
    - Symbols

    1. Basic Latin (U+0000 - U+007F):   
    Decimal Range: 0 - 127
    This is essentially the standard ASCII character set. It includes uppercase and lowercase English letters (A-Z, a-z), digits (0-9), basic punctuation marks, and control characters. This range is fundamental and universally supported.   

    2. Latin-1 Supplement (U+0080 - U+00FF):
    Decimal Range: 128 - 255
    This range extends the Basic Latin set with characters used in many Western European languages. It includes accented letters (like é, à, ü), additional punctuation, currency symbols (like €, £, ¥), and other symbols.   

    3. Latin Extended-A (U+0100 - U+017F) and Latin Extended-B (U+0180 - U+024F):
    Decimal Ranges: 256 - 383 and 384 - 591
    These ranges contain more accented Latin letters and diacritics used in a wider variety of European languages, including Central and Eastern European languages.

    4. General Punctuation (U+2000 - U+206F):
    Decimal Range: 8192 - 8303
    This block contains a variety of punctuation marks, spaces of different widths, and other symbols related to text formatting. You'll find characters like em dash (—), en dash (–), non-breaking space, and various quotation marks here.

    5. Currency Symbols (U+20A0 - U+20CF):
    Decimal Range: 8352 - 8399
    This block includes a wide range of currency symbols from around the world.

    6. Letterlike Symbols (U+2100 - U+214F):
    Decimal Range: 8448 - 8527
    This block contains symbols that look like letters but have special meanings, such as the copyright symbol (©), the registered trademark symbol (®), and various mathematical and technical symbols that resemble letters.

    7. Arrows (U+2190 - U+21FF):
    Decimal Range: 8592 - 8703
    A collection of various arrow symbols.

    8. Mathematical Operators (U+2200 - U+22FF):
    Decimal Range: 8704 - 8959
    Contains a wide array of mathematical symbols.

.PARAMETER Code
    Specifies one or more Unicode code points to convert to characters; accepts integers in the range 1 to 1114111

.PARAMETER Char
    Specifies one or more characters to convert to their Unicode code points (each character must be a single character string)
.PARAMETER IncludeNonPrintable
    Includes non-printable characters in the output (non-printable characters are suppressed by default)

.EXAMPLE
    # Retrieve characters from the "Arrows" Unicode range
    PS C:\> Get-PKSpecialChar -CodeType Arrows

.EXAMPLE
    # Convert a Unicode code point to its corresponding character
    PS C:\> Get-PKSpecialChar -Code 65

.EXAMPLE
    # Convert a character to its corresponding Unicode code point
    Get-PKSpecialChar -Char "A"

.EXAMPLE
    # Retrieve characters from multiple Unicode ranges
    PS C:\> Get-PKSpecialChar -CodeType BasicLatin, Currency

.EXAMPLE
    # Include non-printable characters in the output (suppressed by default)
    PS C:\> Get-PKSpecialChar -CodeType Punctuation -IncludeNonPrintable


#>
[CmdletBinding(DefaultParameterSetName = "Code")]
Param(

    [Parameter(
        ParameterSetName = "Code",
        Position = 0,
        ValueFromPipeline,
        HelpMessage = "Unicode decimal code point to convert to character"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateRange(1,1114111)]
    [Alias("Codepoint","Code")]
    [int[]]$Decimal,

    [Parameter(
        ParameterSetName = "Name",
        HelpMessage = "Character name (partial match)"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$Name,

    [Parameter(
        ParameterSetName = "Range",
        HelpMessage = "Named sets of common unicode characters (see parameter help for details)"
    )]
    [ValidateSet('Arrows','BasicLatin','Currency','LatinExtendedA','LatinExtendedB','LatinSupplement','Mathematical','Punctuation','Symbols')]
    [Alias("Set","CharSet")]
    [string[]]$CharacterSet,

    [Parameter(
        ParameterSetName = "Char",
        Position = 0,
        HelpMessage = "Unicode character to look up"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateLength(1,1)]
    [string[]]$Char,

    [Parameter(
        HelpMessage = "Try to get names via API call to unicode.org"
    )]
    [switch]$GetNames,

    [Parameter(
        HelpMessage = "Include non-printable characters in the output"
    )]
    [switch]$IncludeNonPrintable

)
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
    Where-Object { Test-Path variable:$_ } | Foreach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    $CurrentParams.Add("PipelineInput", $PipelineInput)
    $CurrentParams.Add("ParameterSetName", $Source)
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Lookup tables for code ranges & categories for printable/nonprintable characters
    $RangeLookup = @{
        BasicLatin            = 1..127
        LatinSupplement       = 128..255
        LatinExtendedA        = 256..383
        LatinExtendedB        = 384..591
        Punctuation           = 8192..8303
        Currency              = 8352..8399
        Symbols               = 8448..8527
        Arrows                = 8592..8703
        Mathematical          = 8704..8959
    }

    $VisibleCategories = @(
        [System.Globalization.UnicodeCategory]::UppercaseLetter
        [System.Globalization.UnicodeCategory]::LowercaseLetter
        [System.Globalization.UnicodeCategory]::TitlecaseLetter
        [System.Globalization.UnicodeCategory]::ModifierLetter
        [System.Globalization.UnicodeCategory]::OtherLetter
        [System.Globalization.UnicodeCategory]::NonSpacingMark
        [System.Globalization.UnicodeCategory]::SpacingCombiningMark
        [System.Globalization.UnicodeCategory]::EnclosingMark
        [System.Globalization.UnicodeCategory]::DecimalDigitNumber
        [System.Globalization.UnicodeCategory]::LetterNumber
        [System.Globalization.UnicodeCategory]::OtherNumber
        [System.Globalization.UnicodeCategory]::SpaceSeparator
        [System.Globalization.UnicodeCategory]::LineSeparator
        [System.Globalization.UnicodeCategory]::ParagraphSeparator
        [System.Globalization.UnicodeCategory]::DashPunctuation
        [System.Globalization.UnicodeCategory]::OpenPunctuation
        [System.Globalization.UnicodeCategory]::ClosePunctuation
        [System.Globalization.UnicodeCategory]::InitialQuotePunctuation
        [System.Globalization.UnicodeCategory]::FinalQuotePunctuation
        [System.Globalization.UnicodeCategory]::ConnectorPunctuation
        [System.Globalization.UnicodeCategory]::OtherPunctuation
        [System.Globalization.UnicodeCategory]::MathSymbol
        [System.Globalization.UnicodeCategory]::CurrencySymbol
        [System.Globalization.UnicodeCategory]::ModifierSymbol
        [System.Globalization.UnicodeCategory]::OtherSymbol
        [System.Globalization.UnicodeCategory]::PrivateUse
    )

    $InvisibleCodePoints = @(
        8192..8202 # Various space widths
        8232       # Line Separator (can appear blank)
        8233       # Paragraph Separator (can appear blank)
        8239       # Narrow No-Break Space
        8287       # Ideographic Space
    )
    #endregion

    If ($GetNames.IsPresent) {
        $Msg = "Issuing API call to unicode.org to create lookup table for character names... please be patient"
        Write-Verbose $Msg
        Write-Progress -Activity "Get special characters" -CurrentOperation $Msg
        Function _FormatCharName { # The names are all upper case and it's annoying
            param([Parameter(Mandatory, ValueFromPipeline)][string]$InputString)
            $First = $InputString.Substring(0, 1).ToUpper() # Get the first character
            If ($InputString -match '(.+)(\b\w)$') { # Check if the last character is a standalone alphanumeric
                $Middle = $Matches[1].Substring(1).ToLower() 
                $Last = $Matches[2].ToUpper()
                "$First$Middle`'$Last`'"
            }
            Else {
                $Last = $InputString.Substring(1).ToLower() 
                "$First$Last"
            }
        }

        Try {
            # We'll get the content of the unicodedata file, parse it and create a quick hash table. Then we'll create a custom object with the code (which is hex), convert it to decimal, the name, and the character set.
            $URI = "http://www.unicode.org/Public/UNIDATA/UnicodeData.txt"
            $Request = Invoke-WebRequest $URI -UseBasicParsing -ErrorAction Stop -Verbose:$False
            $NameHash = @{}
            ($Request.content.Trim()  | Where-Object {$_ -ne ""}) | ConvertFrom-CSV -Delimiter ';' -header "CodePoint","Name"  | Foreach-Object {$NameHash[$_.CodePoint] = $_.Name} 
            $UnicodeInfo = $NameHash.GetEnumerator() | Sort-Object Key | Select-Object @{N="Set";E={}},@{N="Hex";E={$_.Key}},@{N="Decimal";E={[int]("0x" + $_.Key)}},@{N="Name";E={$_.Value}} 
            $Select = "Set","Decimal","Hex","Name","Char","Syntax"
            
        }
        Catch {
            $Msg = "Failed to issue API call to unicode.org! $($_.Exception.Message)"
            Write-Error $Msg
            Break
        }
        Finally {Write-Progress -Activity * -Completed}
    }
    Else {
        $Select = "Set","Decimal","Char","Syntax"
    }
    
    $Hidden = @()
    $Output = @()

    $Activity = "Get special / unicode characters"
    Write-Verbose "[BEGIN: $ScriptName] $Activity"
}
Process {
    Switch ($Source) {
        "Range" {
            Foreach ($Set in $CharacterSet) {
                $Range = $RangeLookup[$Set]
                $Msg = "Processing '$Set' (unicode $($Range[0].ToString())-$($Range[-1].ToString()))"
                Write-Verbose $Msg
                Foreach ($C in $Range) {
                    $Lookup = [char]$C
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status "Processing $C" -PercentComplete (($C - $Range[0]) / ($Range[-1] - $Range[0]) * 100)
                    $UnicodeCategory = [System.Char]::GetUnicodeCategory($Lookup) # We don't want to output invisible stuff by default
                    If ((($VisibleCategories -contains $UnicodeCategory) -and ($InvisibleCodePoints -notcontains $C)) -or ($IncludeNonPrintable.IsPresent)) {
                        If ($GetNames.IsPresent) {$InfoLookup = $UnicodeInfo | Where-Object {$_.Decimal -eq $C}}
                        $Output += [PSCustomObject]@{
                            Set     = $Set
                            Decimal = "$("{0:D3}" -f $C)"
                            Hex     =  $Infolookup.Hex ?? $Null
                            Char    = $Lookup
                            Name    = $Infolookup.Name ?? $Null
                            Syntax  = "[char]$C"
                        }
                    }
                    Else {$Hidden += $C}
                }
            }
        } # end range
        "Code" {
            $Total = $Decimal.Count
            $Current = 0
            Foreach ($C in $Decimal) {
                $Current ++
                $Lookup = [char]$C
                $UnicodeCategory = [System.Char]::GetUnicodeCategory($Lookup) # We don't want to output invisible stuff by default
                If ((($VisibleCategories -contains $UnicodeCategory) -and ($InvisibleCodePoints -notcontains $C)) -or ($IncludeNonPrintable.IsPresent)) {
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status "Processing $C" -PercentComplete ($Current/$Total*100)
                    If ($GetNames.IsPresent) {$InfoLookup = $UnicodeInfo | Where-Object {$_.Decimal -eq $C}}
                    Foreach ($Entry in $RangeLookup.GetEnumerator()) {$CurrentRange = $Entry.Value;if ($CurrentRange -contains $C) {$Set = $Entry.Key}}
                    $Output += [PSCustomObject]@{
                        Set     = $Set ?? "(unknown)"
                        Decimal = "$("{0:D3}" -f $C)"
                        Hex     =  $Infolookup.Hex ?? $Null
                        Char    = $Lookup
                        Name    = $Infolookup.Name ?? $Null
                        Syntax  = "[char]$C"
                    }
                }
                Else {$Hidden += $C}
            }
        } # end code
        "Name" {
            $Total = $Name.Count
            $Current = 0
            Foreach ($C in $Name) {
                $Current ++
                $Lookup = $UnicodeInfo | Where-Object {$_.Name -like "$C"}
                Foreach ($L in $Lookup) {
                    Foreach ($Entry in $RangeLookup.GetEnumerator()) {$CurrentRange = $Entry.Value;if ($CurrentRange -contains $C) {$Set = $Entry.Key}}
                    $Output += [PSCustomObject]@{
                        Set     = $Set ?? "(unknown)"
                        Decimal = "$("{0:D3}" -f $C)"
                        Hex     =  $Infolookup.Hex ?? $Null
                        Char    = $Lookup
                        Name    = $Lookup.Name
                        Syntax  = "[char]$C"
                    }
                }

                $Lookup = [char]$C
                $UnicodeCategory = [System.Char]::GetUnicodeCategory($Lookup) # We don't want to output invisible stuff by default
                If ((($VisibleCategories -contains $UnicodeCategory) -and ($InvisibleCodePoints -notcontains $C)) -or ($IncludeNonPrintable.IsPresent)) {
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status "Processing $C" -PercentComplete ($Current/$Total*100)
                    If ($GetNames.IsPresent) {$InfoLookup = $UnicodeInfo | Where-Object {$_.Decimal -eq $C}}
                    Foreach ($Entry in $RangeLookup.GetEnumerator()) {$CurrentRange = $Entry.Value;if ($CurrentRange -contains $C) {$Set = $Entry.Key}}
                    $Output += [PSCustomObject]@{
                        Set     = $Set ?? "(unknown)"
                        Decimal = "$("{0:D3}" -f $C)"
                        Hex     =  $Infolookup.Hex ?? $Null
                        Char    = $Lookup
                        Name    = $Infolookup.Name ?? $Null
                        Syntax  = "[char]$C"
                    }
                }
                Else {$Hidden += $C}
            }
        } # end name
        "Char" {
            $Total = [object[]]$Char.Count
            $Current = 0
            Foreach ($C in $Char) {
        
                Write-Progress -Activity $Activity -CurrentOperation $Msg -Status "Processing char $C" -PercentComplete ($Current++/$Total*100)
                $Lookup = [int]$C[0]
                $UnicodeCategory = [System.Char]::GetUnicodeCategory($Lookup) # We don't want to output invisible stuff by default
                If ((($VisibleCategories -contains $UnicodeCategory) -and ($InvisibleCodePoints -notcontains $C)) -or ($IncludeNonPrintable.IsPresent)) {
                    $InfoLookup = $UnicodeInfo | Where-Object {$_.Decimal -eq $Lookup}
                    Foreach ($Entry in $RangeLookup.GetEnumerator()) {$CurrentRange = $Entry.Value;if ($CurrentRange -contains $C) {$Set = $Entry.Key}}
                    $Output += [PSCustomObject]@{
                        Set     = $Set ?? "(unknown)"
                        Decimal = "$("{0:D3}" -f $Lookup)"
                        Hex     =  $Infolookup.Hex ?? $Null
                        Char    = $C
                        Name    = $Infolookup.Name ?? $Null
                        Syntax  = '[int]"'+"$C"+'"[0]'
                    }
                }
                Else {$Hidden += $C}
            } # end foreach
        } # end char
    } # end switch
}
End {
    Write-Progress -Activity * -Completed
    If ($Hidden.Count -gt 0) {
        Write-Verbose "Suppressing non-printable character values for codes:`n$($Hidden -join(', '))"
    }
    If ($Output.Count -gt 0) {
        Write-Verbose -Message "Returning $($Output.Count) results"
        Write-Output $Output | Where-Object {$_} | Select-Object $Select
    }
    Else {
        Write-Warning "No console-printable characters found in the specified range or code"
    }
}
} # end Get-PKSpecialChar