#requires -version 3
Function Get-PKTempPassword {
<#
.SYNOPSIS
    Generates a custom-length password using alphanumeric/special characters or combo thereof, 
    with option to avoid ambiguous characters

.DESCRIPTION
    Generates a custom-length password using alphanumeric/special characters or combo thereof, 
    with option to avoid ambiguous characters

.NOTES
    Name    : Function_Get-PKTempPassword.ps1 
    Created : 2022-11-14
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2022-11-14 - Created script

.LINK
    https://devblogs.microsoft.com/scripting/generating-a-new-password-with-windows-powershell/

#>
[CmdletBinding()]
Param(

    [Parameter(
        HelpMessage = "Password length (between 1 and 255 characters (default is 12)"
    )]
    [ValidateRange(1,255)] 
    [int]$Length = 12,

    [Parameter(
        HelpMessage = "Include alphabet or ASCII characters (default is both)"
    )]
    [ValidateSet("All","Alphabet","ASCII")]
    [string[]]$Type = "All",

    [switch]$AvoidAmbiguous

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName
    }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Ignore = "1","0","O","l"
    
    $Msg = "Generating $Length-character password"
    Switch ($Type) {
        All      {$Msg += " using both alphabetical & ASCII characters"}
        ASCII    {$Msg += " using ASCII characters"}
        Alphabet {$Msg += " using alphabetical characters"}
    }
    If ($AvoidAmbiguous.IsPresent) {$Msg += ", avoiding common ambiguous characters '$($Ignore -join(', '))'"}

    Write-Verbose $Msg    
}
Process {
    
    $SourceData = @()
    If ($Type -match "Alphabet|All") {
        For ($a=65;$a –le 90;$a++) {$SourceData +=,[char][byte]$a } 
    }
    If ($Type -match "ASCII|All") {
        For ($a=33;$a –le 126;$a++) {$SourceData +=,[char][byte]$a } 
    }

    If ($AvoidAmbiguous.IsPresent) {
        $SourceData = $SourceData | Where-Object {$_ -notmatch $($($Ignore.Join('|')))}
    }

    For ($loop=1; $loop –le $length; $loop++) {
        $TempPassword+=($SourceData | Get-Random)
    }

    Write-Output $TempPassword

}
} 

