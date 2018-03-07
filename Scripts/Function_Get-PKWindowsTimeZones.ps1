#Requires -Version 3
Function Get-PKTimeZones {
<#
.SYNOPSIS
    Returns all time zone info

.DESCRIPTION
    Returns all time zone info
    Returns a PSobject

.NOTES        
    Name    : Function_Get-PKTimeZones.ps1
    Created : 2018-03-05
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-03-05 - Created script
#>
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory = $False,
        Position = 0,
        HelpMessage = "Output type (full or basic)"
    )]
    [ValidateSet("Full","Basic")]
    [string]$OutputType = "Basic",

    [Parameter(
        Mandatory = $False,
        Position = 1,
        HelpMessage = "Sort order (ID or offset)"
    )]
    [ValidateSet("ID","Offset")]
    [string]$SortOrder = "Offset"
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
}
Process {
    
    $Msg = "All time zone information available from $Env:ComputerName"

    Switch ($OutputType) {
        Basic {
            $Msg += " (basic output"
            $Select = "ID","BaseUtcOffset"
        }
        Full {
            $Msg += " (full output"
            $Select = "ID","DisplayName","StandardName","DaylightName","BaseUtcOffset","SupportsDaylightSavingTime"
        }
    }

    Switch ($SortOrder) {
        ID {
            $Msg += ", sorted by ID)"
            $Sort = "ID"
        }
        Offset {
            $Msg += ", sorted by UTC offset)"
            $Sort = "BaseUTCOffset"
        }
    }
    Write-Verbose $Msg

    $AllZones = [System.TimeZoneInfo]::GetSystemTimeZones()
    Return $AllZones | Select $Select | Sort $Sort


}
} #end Get-PKTimeZones