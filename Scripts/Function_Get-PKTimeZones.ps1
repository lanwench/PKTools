#Requires -Version 3
Function Get-PKTimeZones {
<#
.SYNOPSIS
    Retrieves and displays information about system time zones using [System.TimeZoneInfo]::GetSystemTimeZones() 

.DESCRIPTION
    Retrieves and displays information about system time zones using [System.TimeZoneInfo]::GetSystemTimeZones() 
    The output can be customized to show either detailed or minimal information based on the '-Mini' parameter

.NOTES
    Name    : Function_Get-PKTimeZones.ps1
    Created : 2018-03-05
    Author  : Paula Kingsley
    Version : 02.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-03-05 - Created script
        v02.00.0000 - 2025-04-24 - Rewritten because (cough) I learned a lot since 2018; and moved into PKTools module.

.PARAMETER Mini
    If specified, returns only minimal information, including the time zone ID and short/display offset.

.OUTPUTS
    System.Object
    The function outputs a collection of objects containing time zone information. The available properties of the 
    output object includes:
        ~ ID: The identifier of the time zone
        ~ DisplayName: The display name of the time zone
        ~ ShortOffset: A short representation of the UTC offset
        ~ BaseUTCOffset: The base UTC offset of the time zone
        ~ OffsetSeconds: The total offset in seconds
        ~ SupportsDST: Indicates whether the time zone supports daylight saving time
        ~ CurrentTime: The current time in the time zone

.EXAMPLE
    PS C:\> Get-PKTimeZones
    Retrieve detailed time zone information, including ID, display name, offset, and current time

        Id            : Dateline Standard Time
        DisplayName   : (UTC-12:00) International Date Line West
        ShortOffset   : (UTC-12:00)
        BaseUtcOffset : -12:00:00
        OffsetSeconds : -43200
        SupportsDST   : False
        CurrentTime   : 4/24/2025 9:37:32 AM

        Id            : UTC-11
        DisplayName   : (UTC-11:00) Coordinated Universal Time-11
        ShortOffset   : (UTC-11:00)
        BaseUtcOffset : -11:00:00
        OffsetSeconds : -39600
        SupportsDST   : False
        CurrentTime   : 4/24/2025 10:37:32 AM

        Id            : Aleutian Standard Time
        DisplayName   : (UTC-10:00) Aleutian Islands
        ShortOffset   : (UTC-10:00)
        BaseUtcOffset : -10:00:00
        OffsetSeconds : -36000
        SupportsDST   : True
        CurrentTime   : 4/24/2025 11:37:32 AM
        
        <snip>

.EXAMPLE
    PS C:\>Get-PKTimeZones -Mini
    Retrieve minimal time zone information with only ID and short offset

        Id                             ShortOffset
        --                             -----------
        Dateline Standard Time         (UTC-12:00)
        UTC-11                         (UTC-11:00)
        Aleutian Standard Time         (UTC-10:00)
        Hawaiian Standard Time         (UTC-10:00)
        Marquesas Standard Time        (UTC-09:30)
        Alaskan Standard Time          (UTC-09:00)
        UTC-09                         (UTC-09:00)
        Pacific Standard Time (Mexico) (UTC-08:00)
        UTC-08                         (UTC-08:00)
        Pacific Standard Time          (UTC-08:00)
        
        <snip>


#>

[CmdletBinding()]
Param(
    [Parameter(
        HelpMessage = "Display only minimal output (ID and short offset)"
    )]
    [switch]$Mini
)
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"

    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Msg = "Getting time zone information"
    If ($Mini) {
        $Msg += " (basic output only)"
        $Select = "ID","ShortOffset"
    }
    Else {
        $Msg += " (detailed output)"
        $Select = "*"
    }
    Write-Verbose "[BEGIN: $ScriptName] $Msg"
}
Process {
    
    Try {
        $TimeTable = [System.TimeZoneInfo]::GetSystemTimeZones() | Sort-Object BaseUTCOffSet |
            Select-Object ID,
            DisplayName,
            @{N="ShortOffset";E={
                If ($_.ID -eq "UTC") {"(UTC+00:00)"}
                Else {
                    if ($_.DisplayName -match '\(UTC[+-]\d{2}:\d{2}\)') {$Matches[0]} else {$null}}
                }
            },
            BaseUTCOffset,
            @{N="OffsetSeconds";E={$_.BaseUtcOffset.TotalSeconds}},
            @{N="SupportsDST";E={$_.SupportsDaylightSavingTime}},
            @{N="CurrentTime";E={(Get-Date -AsUTC).AddSeconds($_.BaseUtcOffset.TotalSeconds)}}

        $TimeTable | Select-Object $Select 
    }
    Catch {Throw $_.Exception.Message}
}
End {
    Write-Verbose "[END: $ScriptName] Operation complete"
}
} #end Get-PKTimeZones
