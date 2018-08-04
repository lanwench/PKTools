#requires -Version 3
Function New-PKRandomPassword {
<#
.SYNOPSIS
    Generates a random password string, with option to select length

.DESCRIPTION
    Generates a random password string, with option to select length
    Outputs a string
    
.NOTES
    Name    : Function_New-PKRandomPassword.ps1 
    Created : 2018-04-02
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2018-04-02 - Created script based on Steve König's original (see link)

.LINK
    http://activedirectoryfaq.com/2017/08/creating-individual-random-passwords/

.EXAMPLE
    PS C:\> New-PKRandomPassword -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key           Value               
        ---           -----               
        Verbose       True                
        Length        10                  
        Debug         Active Directory    
        ScriptName    New-PKRandomPassword
        ScriptVersion 1.0.0               

        VERBOSE: Generate 10-character password
        "ptukk2iSlc

#>
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory = $False,
        HelpMessage = "Length"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateRange(5,255)]
    [int]$Length = 10
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"
    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Functions

    function Get-RandomCharacters($length, $characters) {
        $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
        $private:ofs=""
        return [String]$characters[$random]
    }
 
    function Scramble-String([string]$inputString){     
        $characterArray = $inputString.ToCharArray()   
        $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
        $outputString = -join $scrambledStringArray
        return $outputString 
    }
 
    #endregion Functions

    #region Strings of characters

    $String_lc   = "abcdefghiklmnoprstuvwxyz"
    $String_uc   = "ABCDEFGHKLMNOPRSTUVWXYZ"
    $String_int  = "1234567890"
    $String_spec = "@#$%^&*–_+=[]{}|\:',?/~""();.£"

    #endregion Strings of characters

    $Msg = "Generate $Length-character password"
    Write-Verbose $Msg
}
Process {
    
    $password  = Get-RandomCharacters -length 8 -characters $String_lc
    $password += Get-RandomCharacters -length 1 -characters $String_uc
    $password += Get-RandomCharacters -length 1 -characters $String_int
    $password += Get-RandomCharacters -length 1 -characters $String_spec

    $password = Scramble-String $password
 
    Write-Output $password

}

} #end New-PKRandomPassword

