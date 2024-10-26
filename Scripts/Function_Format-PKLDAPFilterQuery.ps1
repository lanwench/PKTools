#requires -version 4
Function Format-PKLDAPFilter {
<# 
.SYNOPSIS
    Formats an LDAP filter string with line breaks and indentation, outputting as a more visually readable string

.DESCRIPTION
    Formats an LDAP filter string with line breaks and indentation, outputting as a more visually readable string
    To revert to valid ldap filter, use $filter -replace '\n\s*',''
    Accepts pipeline input

.NOTES
    Name    : Function_Format-PKLDAPFilter.ps1
    Created : 2024-08-22
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2024-08-22 - Created script based on Simon Wahlin's original (thanks, wayback machine!)

.LINK
    https://blog.simonw.se/working-with-complex-ldap-filters-in-powershell/ 

.LINK
    https://gallery.technet.microsoft.com/Rewrite-an-ldap-filter-in-d086f731/file/116883/1/Show-SWLdapFilter.ps1 

.PARAMETER ldapFilter
    One or more LDAP filter strings to format

.EXAMPLE
    PS C:\> $filter = '(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(|(accountExpires=0)(accountExpires=9223372036854775807))(userAccountControl:1.2.840.113556.1.4.803:=65536))'
    PS C:\> ConvertFrom-PKLDAPFilter -ldapFilter $filter
        (&
            (objectCategory=person)
            (objectClass=user)
            (!
                (userAccountControl:1.2.840.113556.1.4.803:=2)
            )
            (|
                (accountExpires=0)
                (accountExpires=9223372036854775807)
            )
            (userAccountControl:1.2.840.113556.1.4.803:=65536)
        )

#>
[Cmdletbinding()]
Param (
    [Parameter(
        Mandatory,
        Position = 0,
        ValueFromPipeline
    )]
    [String[]] $ldapFilter
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $stringBuilder = New-Object System.Text.StringBuilder
    #$Tab = "    "

    # internal function
    function New-IndentedText ($Text,$NumIndents) {
        "`n$("    " * $NumIndents)$Text"
    }

    function Get-LogicalOperatorDescription {
        param ([Parameter(Mandatory,ValueFromPipeline,Position=0)][string]$operator)
        switch -regex ($operator) {
            '&' { "AND (Must match all)" }
            '|' { "OR (Must match any)" }
            '!' { "NOT (Must not match)" }
            default { return "Unknown logical operator" }
        }
    }

    $Msg = "[BEGIN: $ScriptName] Formatting LDAP filter strings"
    Write-Verbose $Msg
}
Process {

Foreach ($String in $ldapFilter){
    $null = $stringBuilder.Clear()
    $Indent = 0
    $null = $stringBuilder.Append($String[0])
    
    for( $i=1; $i -lt $String.Length; $i++ ){
        $Char = $String[$i]
        switch -Regex ($String.Substring( ($i-1),2 )) {
            "\([|&!]" { # match an open parenthesis followed by any one of the characters |, &, or !
                $Indent++
                $Operator = $Char | Get-LogicalOperatorDescription
                #$Char = "$Char [$Operator]"
                $null = $stringBuilder.Append("$Char [$Operator]")
                $null = $stringBuilder.Append($(New-IndentedText "" $Indent))
            }
            "\)\(" { # match closing parenthesis immediately followed by an open parenthesis ....)(
                $null = $stringBuilder.Append($(New-IndentedText $Char $Indent))
            }
            "\)\)" { #  match two consecutive closing parentheses .... ))
                $Indent--
                $null = $stringBuilder.Append($(New-IndentedText $Char $Indent))
            }
            default{
                $null = $stringBuilder.Append($Char)
            }
        }
    }
    if($Indent -lt 0){
        Throw "Invalid ldapFilter!"
    }
    $stringBuilder.ToString()
}
}
End {
    $Msg = "[END: $ScriptName]"
    Write-Verbose $Msg

    $Null | Remove-Variable -Name StringBuilder -ErrorAction SilentlyContinue
}
} #end Format-PKLDAPFilter 

