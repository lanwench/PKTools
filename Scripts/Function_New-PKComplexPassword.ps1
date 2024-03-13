#requires -version 5
function New-PKComplexPassword {
    <#
.SYNOPSIS
    Uses Get-Random and defined character sets to generate a password between 10 and 265 characters, with option to return secure string or plain text

.DESCRIPTION
    Uses Get-Random and defined character sets to generate a password between 10 and 265 characters, with options to return secure string or plain text
    Will include mixed case alpha and no more (or less) than 2 special characters
    Returns a string or securestring object
    
.NOTES        
    Name    : Function_New-PKComplexPassword.ps1
    Created : 2024-02-12
    Author  : Paula Kingsley
    Version : 01.00.1000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2024-03-12 - Created to replace New-PKRandomPassword, based on Yann Normand's original (see link)


.PARAMETER Length
    Password length (integer between 10 and 265; default is 14)

.PARAMETER Number
    Number of passwords to return, between 1 and 30 (default is 1)

.PARAMETER OutputType
    Return password as secure string (default) or plain text (not recommended for most purposes!)

.OUTPUTS
    Secure string or plain text

.LINK
    https://dev.to/onlyann/user-password-generation-in-powershell-core-1g91

.EXAMPLE
    PS C:\>  New-PKComplexPassword -Verbose

        VERBOSE: PSBoundParameters: 

        Key              Value
        ---              -----
        Verbose          True
        Length           14
        Number           1
        PlainText        False
        ScriptName       New-PKComplexPassword
        ParameterSetName SecureString
        ScriptVersion    1.1.0

        VERBOSE: Creating 1 14-character complex password(s) in secure string format

        VERBOSE: [1 of 1]
        System.Security.SecureString

.EXAMPLE    
    PS C:\> New-PKComplexPassword -PlainText -Length 12 -AvoidAmbiguous -NumSpecialChars 3 -Number 4 -Verbose 
        VERBOSE: PSBoundParameters: 

        Key              Value
        ---              -----
        PlainText        True
        Length           12
        AvoidAmbiguous   True
        NumSpecialChars  3
        Number           4
        Verbose          True
        ScriptName       New-PKComplexPassword
        ParameterSetName PlainText
        ScriptVersion    1.0.0

        VERBOSE: Creating 4 12-character complex password(s) with 3 special characters, in plain text format
        WARNING: Plain text is insecure; please use this only for testing purposes!

        VERBOSE: [1 of 4]
        MkZEqz!%U#Ig5
        VERBOSE: [2 of 4]
        vFmzb&K@7&YIW
        VERBOSE: [3 of 4]
        i!C2xHu*mfjq*
        VERBOSE: [4 of 4]
        r@zi4HHe#d@HD

    #>
    [CmdletBinding(DefaultParameterSetName = "SecureString")]
    param(
        [Parameter(
            Position = 0,
            HelpMessage = "Password length (integer between 10 and 265; default is 14)"
        )]
        [ValidateRange(10, 256)]
        [int]$Length = 14,

        [Parameter(
            HelpMessage = "Number of passwords to return, between 1 and 30 (default is 1)"
        )]
        [ValidateRange(1, 30)]
        [int]$Number = 1,

        [Parameter(
            HelpMessage = "Avoid ambiguous characters such as 1, l, 0, O"
        )]
        [switch]$AvoidAmbiguous,

        [Parameter(
            HelpMessage = "Exact number of special characters to include, between 1 and 8 (default is 2)"
        )]
        [ValidateRange(1, 8)]
        [int]$NumSpecialChars = 2,

        [Parameter(
            ParameterSetName = "PlainText",
            HelpMessage = "Return password in plain text instead of default secure string (use this carefully as it is not secure)"
        )]
        [switch]$PlainText
    )
    Begin {

        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # Show our settings
        $ScriptName = $MyInvocation.MyCommand.Name
        $Source = $PSCmdlet.ParameterSetName
        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ParameterSetName", $Source)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        $Msg = "Generating $Number complex $Length-character $(switch ($Number) {1 {"password"};default {"passwords"}}) with $NumSpecialChars special $(switch ($NumSpecialChars) {1 {"character"};default {"characters"}})"
        Switch ($Source) {
            SecureString {
                $Msg += ", in secure string format`n"
                Write-Verbose $Msg
            }
            PlainText {
                $Msg += ", in plain text format"
                Write-Verbose $Msg
                $Msg = "Plain text is insecure; please use this only for testing purposes!`n"
                Write-Warning $Msg
            }
        }           

        # Allowed characters
        $Symbols = '!@#$%^&*'.ToCharArray()
        $EscapedChars = [regex]::escape($Symbols -join(";")).Replace(";","|")
        
        $CharacterList = 'a'..'z' + 'A'..'Z' + '0'..'9' + $symbols
        If ($AvoidAmbiguous.IsPresent) { $Ambiguous = "1", "0", "O", "l" }
        Else { $Ambiguous = @() }
        [int]$MaxTries = 100
        
    }
    Process {

        for ($i = 1; $i -le $Number; $i++) {
            
            $Msg = "[$i of $Number]"
            Write-Verbose $Msg
            Try {
                [int]$Counter = 0
                do {
                    $password = -join (0..$length | Foreach-Object { $characterList | Where-Object { $Ambiguous -notcontains $_ } | Get-Random -ErrorAction Stop })
                    $Counter ++
                    [int]$hasLowerChar = $password -cmatch '[a-z]' # 1
                    [int]$hasUpperChar = $password -cmatch '[A-Z]' # 2
                    [int]$hasDigit = $password -match '[0-9]'  # 3
                    [int]$hasSymbol = $password.IndexOfAny($symbols) -ne -1 # 4
                    [int]$numSymbol = (([regex]::Matches($Password,$EscapedChars)).Count -eq $NumSpecialChars) # 5
                }
                until ((($hasLowerChar + $hasUpperChar + $hasDigit + $hasSymbol + $numSymbol) -eq 5) -or ($Counter -eq $MaxTries))
                
                If ($Counter -eq $MaxTries) {
                    $Msg = "Unable to create acceptable password after ($Counter) attempts, sorry!"
                    Write-Warning $Msg
                }
                Else {
                    Switch ($Source) {
                        SecureString {
                            $Password | ConvertTo-SecureString -AsPlainText
                        }
                        PlainText {
                            $Password
                        }
                    }
                }
            }
            Catch {
                Throw $_.Exception.Message
            }
        }
    }
} #end function

$Null = New-Alias New-PKRandomPassword -Value New-PKComplexPassword -Force -Description "Replacing old function, 2024-03-12"
