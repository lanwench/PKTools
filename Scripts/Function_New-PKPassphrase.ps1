#requires -version 4
function New-PKPassphrase {
<#
.SYNOPSIS 
    Generates one or more passphrases of English or Lorem Ipsum, using REST API calls. Allows selection of word count, integer count, and separator.

.DESCRIPTION
    Generates one or more passphrases of English or Lorem Ipsum, using REST API calls. Allows selection of word count, integer count, and separator.
    Returns one or more strings

    NEED TO ADD/FIX:
    Total passphrase length param
    Empty lines not being removed

.OUTPUTS
    String

.NOTES
    Name    : Function_New-PKPassphrase.ps1
    Created : 2024-04-03
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2024-04-03 - Created script

.LINK
    https://scriptposse.com/2020/04/24/random-word-function/

.LINK
    https://publicapis.io/loripsum-api

.PARAMETER Dictionary
    English or Lorem Ipsum words (default is English)

.PARAMETER Count
    Number of passphrases to generate, between 1-100 (default is 1)

.PARAMETER NumWords
    Number of unique words for passphrase, between 4-8 (default is 4)

.PARAMETER NumIntegers
    Number of unique integers for passphrase, between 1-5 (default is 1)

.PARAMETER Separator
    Separator character: hyphen, space, carat, or underscore (default is hyphen)

.EXAMPLE
    PS C:\ New-PKPassphrase -Verbose

        VERBOSE: Generating passphrase consisting of 5 English words and 1 integer(s), separated by hyphens
        VERBOSE: Passphrase 1 of 1...
        hexane-taunt-galleon-Possets-8
        
.EXAMPLE
    PS C:\> New-PKPassphrase -Count 5 -Dictionary LoremIpsum -NumWords 5 -NumIntegers 1

        quos-videri-illud-verum-9-Causa                                                                                         
        Certe-2-iure-rebus-boni-caius
        nam-dolor-2-inquam-Elit-habere
        rerum-Sunt-croesi-multa-2-non
        9-videmus-magis-horum-Res-vobis
#>
    param (
        [Parameter(
            HelpMessage = "English or Lorem Ipsum words (default is English)"
        )]
        [ValidateSet("English","LoremIpsum")]
        [string]$Dictionary = "English",

        [Parameter(
            HelpMessage = "Number of passphrases to generate, between 1-100 (default is 1)"
        )]
        [ValidateRange(1,100)]
        [int]$Count = 1,

        [Parameter(
            HelpMessage = "Number of unique words for passphrase, between 4-8 (default is 4)"
        )]
        [ValidateRange(4,8)]
        [int]$NumWords = 4,

        [Parameter(
            HelpMessage = "Number of unique integers for passphrase, between 1-5 (default is 1)"
        )]
        [ValidateRange(1,5)]
        [int]$NumIntegers = 1,

        [Parameter(
            HelpMessage = "Separator character: hyphen, space, carat, or underscore (default is hyphen)"
        )]
        [ValidateSet("Hyphen", "Space", "Carat","Underscore")]
        [string]$Separator = "Hyphen"
    )
    Begin {

        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # Show our settings
        $ScriptName = $MyInvocation.MyCommand.Name

        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path variable:$_ } | ForEach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        $Msg = "Generating passphrase consisting of $NumWords $Dictionary words and $NumIntegers integer(s), separated by $($Separator.tolower())s"
        Write-Verbose $Msg

        Switch ($Dictionary) {
            English {
                $URI = "https://raw.githubusercontent.com/RazorSh4rk/random-word-api/master/words.json"
                $AllWords = (Invoke-WebRequest -Uri $URI -UseBasicParsing -Method Get -Verbose:$False -ErrorAction SilentlyContinue).content | ConvertFrom-JSON | Where-Object {$_.Length -lt 8}
            }
            LoremIpsum {
                $URI = "https://loripsum.net/api?type=short"
                $AllWords = (((Invoke-WebRequest -Uri $URI -UseBasicParsing -Method Get -Verbose:$False -ErrorAction SilentlyContinue).Content -replace "[^\w\s]+") -split(" ")).Trim().ToLower() | 
                    Select-Object -Unique -CaseInsensitive | Where-Object {$_.Length -gt 2 -and $_.length -lt 8}
            }
        }
        
        $Sep = Switch ($Separator) {
            Hyphen {"-"}
            Space  {" "}
            Carat  {"^"}
            Underscore {"_"}
        }
    }
    Process {
        
        For ($i=0;$i -lt $Count;$i ++) {
            $Msg = "Passphrase $([int]$i +1) of $Count..."
            Write-Verbose $Msg
            Try {
                $IntArr = (Get-Random -InputObject 2,3,4,5,6,7,8,9 -count $NumIntegers)
                $WordArr = $AllWords | Get-Random -Count $NumWords
                $Upper = (Get-Culture).TextInfo.ToTitleCase(($WordArr | Get-Random -count 1))        
                $PassArr = @(($WordArr -replace($Upper,$Upper)) + $IntArr)
                $PassPhrase = ($PassArr | Sort-Object {Get-Random}) -join($Sep)
                Write-Output $Passphrase
            }
            Catch {
                Write-Warning $_.Exception.Message
            }
        }
    }
} #end function
