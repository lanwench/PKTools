#requires -Version 3
Function ConvertTo-PKRegexArray {
<#
.Synopsis
    Converts a simple array of strings to a regular expression with escaped characters

.DESCRIPTION
    Converts a simple array of strings to a regular expression with escaped characters
    Created because I always forget the syntax and frequently need to compare
    strings / partial matches in an array
    Returns a string

.NOTES
    Name    : Function_ConvertTo-PKRegexArray
    Created : 2017-11-13
    Version : 01.00.0000
    Author  : Paula Kingsley
    History:
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-11-13 - Created script

.PARAMETER InputObject
    Array object of strings to convert to regular expression with escaped characters

        
.EXAMPLE
    PS C:\> $Arr = "apple","ba*nana","four tomatoes","water-melon"
    
    PS C:\> ConvertTo-PKRegexArray $Arr
        apple|ba\*nana|four\ tomatoes|water-melon
    
    PS C:\> (ConvertTo-PKRegexArray $Arr) -match "tomato"
        True
    
#>
[Cmdletbinding()]
Param(
    [Parameter(
        Mandatory = $True,
        ValueFromPipeline = $True
    )]
    [ValidateNotNullOrEmpty()]
    [array]$InputObject)

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

    # Preference
    $ErrorActionPreference = "Stop"

}

Process {
    
    ($InputObject | 
        Foreach-Object { 
            [regex]::escape($_) } ) -join "|"

}

} #end ConvertTo-PKRegexArray