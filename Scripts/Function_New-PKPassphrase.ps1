#requires -version 4
function New-PKPassphrase {
    <#
    .SYNOPSIS 
        Uses REST API calls to generates one or more passphrases of English or Lorem Ipsum, 
        with the option to pick the count, the word count, the separator character, and number of integers
    
    .DESCRIPTION
        Uses REST API calls to generates one or more passphrases of English or Lorem Ipsum, 
        with the option to pick the count, the word count, the separator character, and number of integers
        Available output formats are string, securestring (string should be used only for testing!
    
    .OUTPUTS
        String
    
    .OUTPUTS
        SecureString
    .NOTES
        Name    : Function_New-PKPassphrase.ps1
        Created : 2024-04-03
        Author  : Paula Kingsley
        Version : 03.00.0000
        History :
        
            ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **
    
            v01.00.0000 - 2024-04-03 - Created script
            v02.00.0000 - 2024-04-26 - Learned stuff, fixed stuff, added stuff
            v03.00.0000 - 2025-03-21 - Replaced loripsum.net api with https://fakeapi.net/lorem
    
    .LINK
        https://scriptposse.com/2020/04/24/random-word-function/
    
    .LINK
        https://publicapis.io/loripsum-api
    
    .PARAMETER Dictionary
        English or Lorem Ipsum words (default is English)
    
    .PARAMETER Count
    
    .PARAMETER PassphraseLength 
        Number of passphrases to generate, between 1-100 (default is 1)
    
    .PARAMETER NumWords
        Number of unique words for passphrase, between 4-8 (default is 4)
    
    .PARAMETER NumIntegers
        Number of unique integers for passphrase, between 1-5 (default is 1)
    
    .PARAMETER Separator
        Separator character: hyphen, space, carat, or underscore (default is hyphen)
    
    .PARAMETER ReturnAs
        Return output as string or secure string
    
    .EXAMPLE
        PS C:\> New-PKPassphrase -ReturnAs String -Verbose 
    
            VERBOSE: PSBoundParameters: 
            Key              Value
            ---              -----
            ReturnAs         String
            Verbose          True
            Dictionary       English
            Count            1
            PassphraseLength 14
            NumWords         4
            NumIntegers      1
            Separator        Hyphen
            ScriptName       New-PKPassphrase
            ScriptVersion    1.0.0
            
            VERBOSE: Generating passphrase consisting of 4 English words and 1 integer(s), separated by hyphen characters, returning in String format
            WARNING: Returning passphrase as string - this is not secure, and should be avoided. Included for testing and entertainment purposes. For production use, consider using -ReturnAs SecureString instead!
            VERBOSE: Passphrase 1 of 1...
            holders-carbos-numbats-7-Untaxed
    
    .EXAMPLE
        PS C:\> New-PKPassphrase -Count 5 -Dictionary LoremIpsum -NumWords 4 -NumIntegers 1 -PassphraseLength 20 -ReturnAs String
    
            WARNING: Returning passphrase as string - this is not secure, and should be avoided. Consider using -ReturnAs SecureString instead!
            essent-7-contra-Quidem-Qui                                                                                              
            elit-iucunde-Est-ambarum-6                                                                                              
            ista-Quo-omnia-3-illi                                                                                                   
            8-nostram-omnia-non-Iussit                                                                                              
            tandem-ipse-Aut-4-omne        
        
    .EXAMPLE
        PS C:\ New-PKPassphrase -Dictionary LoremIpsum -NumWords 4 -NumIntegers 1 -PassphraseLength 20 -ReturnAs SecureString         
            System.Security.SecureString 
            
    .EXAMPLE
        PS C:\> New-PKPassphrase -Count 5 -Dictionary LoremIpsum -NumWords 5 -NumIntegers 1
    
            quos-videri-illud-verum-9-Causa                                                                                         
            Certe-2-iure-rebus-boni-caius
            nam-dolor-2-inquam-Elit-habere
            rerum-Sunt-croesi-multa-2-non
            9-videmus-magis-horum-Res-vobis
    #>
        [CmdletBinding()]
        Param (
            [Parameter(
                HelpMessage = "English or Lorem Ipsum dictionary words (default is English)"
            )]
            [ValidateSet("English","LoremIpsum")]
            [string]$Dictionary = "English",
    
            [Parameter(
                HelpMessage = "Number of passphrases to generate, between 1-100 (default is 1)"
            )]
            [ValidateRange(1,100)]
            [int]$Count = 1,
    
            [Parameter(
                HelpMessage = "Total length of passphrase, including separator characters, between 12-36 characters (default is 14)"
            )]
            [ValidateRange(12,36)]
            [int]$PassphraseLength = 14,
    
            [Parameter(
                HelpMessage = "Number of unique words for passphrase, between 3-8 (default is 4)"
            )]
            [ValidateRange(3,8)]
            [int]$NumWords = 4,
    
            [Parameter(
                HelpMessage = "Number of unique integers for passphrase, between 0-5 (default is 1)"
            )]
            [ValidateRange(0,5)]
            [int]$NumIntegers = 1,
    
            [Parameter(
                HelpMessage = "Separator character: hyphen, space, carat, or underscore (default is hyphen)"
            )]
            [ValidateSet("Hyphen", "Space", "Carat","Underscore")]
            [string]$Separator = "Hyphen",
    
            [Parameter(
                Mandatory,
                HelpMessage = "Return output as string or secure string"
            )]
            [ValidateSet("String","SecureString")]
            [string]$ReturnAs
        )
        Begin {
    
            # Current version (please keep up to date from comment block)
            [version]$Version = "03.00.0000"
    
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
    
            #region Setup
    
            Switch ($Dictionary) {
                English {$URI = "https://raw.githubusercontent.com/RazorSh4rk/random-word-api/master/words.json"}
                LoremIpsum {$URI = "https://fakeapi.net/lorem/50"} # Returns 50 by default but then we select 2-6 char max, so we need extra
            }
    
            Function _GetDictionary{
                Try {
                    Switch ($Dictionary) {
                        English {
                            ((Invoke-WebRequest -Uri $URI -UseBasicParsing -Method Get -Verbose:$False -ErrorAction SilentlyContinue).content | 
                                ConvertFrom-JSON | 
                                    Where-Object {$_.Length -gt 2 -and $_.length -lt 8}) | 
                                        ForEach-Object {$_.Trim().ToLower()} 
                        }
                        LoremIpsum {
                            # This uri will return 50 by default - we can then get the number we want later. But we first want to make sure the strings are between 2 - 8 chars max.
                            ((Invoke-WebRequest -Uri $URI -UseBasicParsing -Method Get -Verbose:$False -ErrorAction SilentlyContinue).Content | ConvertFrom-JSON | Select-Object -ExpandProperty Text) -split(" ") | 
                                Where-Object {$_.Length -gt 2 -and $_.length -lt 8} | 
                                    ForEach-Object {$_.Trim().ToLower()} |  
                                        Select-Object -Unique -CaseInsensitive
                        }
                    }
                }
                Catch {
                    Throw $_.Exception.Message
                }
            }
            
            $SepChar = Switch ($Separator) {
                Hyphen      {"-"}
                Space       {" "}
                Carat       {"^"}
                Underscore  {"_"}
            }
            Function _Shuffle {
                Param (
                    [Parameter(ValueFromPipeline, Position=0)]
                    [string]$String,
                    [string]$Sep = $SepChar
                )
                Do {
                    [object[]]$Arr = $String -split $Sep
                    [object[]]$Shuffled = $Arr | Get-Random -Count $Arr.Count
                    [string]$Out = ($Shuffled -join $Sep)
                } Until ($Out -ne $String)
                Write-Output $Out 
            }
    
            #endregion Setup
    
            $Msg = "Generating passphrase consisting of $NumWords $Dictionary words and $NumIntegers integer(s), separated by $($Separator.tolower()) characters, returning in $ReturnAs format"
            Write-Verbose $Msg
    
            If ($ReturnAs -eq "String") {
                $Msg = "Returning passphrase as string - this is not secure, and should be avoided. Included for testing and entertainment purposes. For production use, consider using -ReturnAs SecureString instead!"
                Write-Warning $Msg
            }
        }
        Process {
    
            For ($i=0;$i -lt $Count;$i ++) {
    
                $Msg = "Passphrase $([int]$i +1) of $Count..."
                Write-Verbose $Msg
                $Passphrase = $Null
                
                Try {
                    $AllWords = _GetDictionary
                    Do {
                        If ($NumIntegers -gt 0) {$IntArr = (Get-Random -InputObject 2,3,4,5,6,7,8,9 -count $NumIntegers)}         
                        $WordArr = $AllWords | Get-Random -Count $NumWords
                        $Upper = (Get-Culture).TextInfo.ToTitleCase(($WordArr | Get-Random -count 1))        
                        $PassArr = @(($WordArr -replace($Upper,$Upper)) + $IntArr)
                        $PassPhrase = ($PassArr | Sort-Object {Get-Random } -Unique) -join($SepChar)
                    } While ($Passphrase.length -lt $PassphraseLength)
    
                    # Shuffle it for good luck
                    $Output = ($Passphrase | _Shuffle)
                    Switch ($ReturnAs) {
                        String {Write-Output $Output}
                        SecureString {ConvertTo-SecureString -String $Output -AsPlainText -Force}
                    }
                }
                Catch {
                    Write-Warning $_.Exception.Message
                }
                Finally {
                    $Passphrase = $Null
                }
            }
        }
    } #end New-PKPassphrase
    