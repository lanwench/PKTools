#requires -Version 3
Function ConvertTo-PKRegex {
<#
.SYNOPSIS
    Escapes characters in one or more strings for nefarious regex purposes

.DESCRIPTION
    Escapes characters in one or more strings for nefarious regex purposes
    Created because I always forget the syntax and frequently need to compare
    strings / partial matches in an array
    If -ReturnString is specified, joins results with | character

.NOTES
    Name    : Function_ConvertTo-PKRegex.ps1
    Created : 2017-11-13
    Version : 02.00.0000
    Author  : Paula Kingsley
    History:
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-11-13 - Created script
        v02.00.0000 - 2023-08-03 - Renamed, added option for output as individual string(s) or single string 

.PARAMETER Text
    One or more strings (or arrays of strings), containing text / characters to escape

.PARAMETER ReturnString
    Return output as a single string of escaped text strings, separated by | character
        
.EXAMPLE
    PS C:\>ConvertTo-PKRegex 'f*teen','$ sign','dollar bill' -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                        
        ---              -----                        
        Verbose          True                         
        Text             {f*teen, $ sign, dollar bill}
        ReturnString     False                        
        ScriptName       ConvertTo-PKRegex            
        ScriptVersion    2.0.0                        
        ParameterSetName Default                      
        PipelineInput    False                        

        VERBOSE: [BEGIN: ConvertTo-PKRegex] Escape characters in one or more text strings
        VERBOSE: Adding 'f*teen'
        VERBOSE: Adding '$ sign'
        VERBOSE: Adding 'dollar bill'
        VERBOSE: Total collection item count : 3
        
        f\*teen
        \$\ sign
        dollar\ bill

        VERBOSE: [END: ConvertTo-PKRegex]

.EXAMPLE
    PS C:\> 'f*teen','$ sign','dollar bill' | ConvertTo-PKRegex -ReturnString -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value            
        ---              -----            
        ReturnString     True             
        Verbose          True             
        Text                              
        ScriptName       ConvertTo-PKRegex
        ScriptVersion    2.0.0            
        ParameterSetName String           
        PipelineInput    True             

        VERBOSE: [BEGIN: ConvertTo-PKRegex] Escape characters in one or more text strings
        VERBOSE: Adding 'f*teen'
        VERBOSE: Adding '$ sign'
        VERBOSE: Adding 'dollar bill'
        VERBOSE: Total collection item count : 3
        
        f\*teen|\$\ sign|dollar\ bill

        VERBOSE: [END: ConvertTo-PKRegex]

.EXAMPLE
    PS C:\> "fourteen or fifteen dollar bills" -match ('f*teen','$ sign','dollar bill' | ConvertTo-PKRegex -ReturnString)
        True

#>
[Cmdletbinding()]
Param(
    [Parameter(
        Position = 0,
        Mandatory = $True,
        ValueFromPipeline = $True,
        HelpMessage = "One or more strings (or arrays of strings), containing text / characters to escape"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$Text,

    [Parameter(
        HelpMessage = "Return output as a single string of escaped-text strings, separated by | character"
    )]
    [switch]$ReturnString
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"

    # How did we get here?
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $ScriptName = $MyInvocation.MyCommand.Name
    $Source = $PSCmdlet.ParameterSetName

    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # We want to collect all the input into a single output string even if it's from the pipeline 
    # Thank you as always, Jeff!
    $Data = [System.Collections.Generic.List[object]]::New()

   $Activity = "Escape characters in one or more text strings"
   If ($ReturnString.IsPresent) {$Activity += ", returning single string separated by | character"}
   Write-Verbose "[BEGIN: $ScriptName] $Activity"

}
Process {
    
    Foreach ($i in $Text) {
        if ($i -is [array]) {
            $Msg = "Adding '$($i -join("', '"))'"
            Write-Verbose $Msg
            $Data.AddRange($i)
        }
        else {
            $Msg = "Adding '$i'"
            Write-Verbose $Msg
            $Data.Add($i)
        }
    }

}
End {
    
    $Msg = "Total collection item count : $($Data.Count)"
    Write-Verbose $Msg

    $Output = ($Data | Foreach-Object {[regex]::escape($_) } )

    If ($ReturnString.IsPresent) {Write-Output ($Output -join "|")}
    Else {Write-Output $Output}
   
    Write-Verbose "[END: $ScriptName]"
}
} #end ConvertTo-PKRegex

$Null = New-Alias ConvertTo-PKRegexArray -Value ConvertTo-PKRegex -Description "Backwards compatibility" -Force

