#Requires -Version 3
Function Get-PKWindowsDateTime {
<#
.SYNOPSIS
    Returns various date / time / time zone settings for a computer

.DESCRIPTION
    Returns various date / time / time zone settings for a computer
    Returns a PSobject

.NOTES        
    Name    : Function_Get-PKWindowsDateTime.ps1
    Created : 2018-03-05
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-03-05 - Created script
#>
[CmdletBinding()]
Param()
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
    
    $Msg = "Current time and timezone settings on $Env:ComputerName"
    Write-Verbose $Msg

    $CurrTime = [System.DateTime]::Now
    $ZoneInfo = [TimeZoneInfo]::Local
    $NTP = & w32tm /query /status

    New-Object PSObject -Property ([ordered] @{
        ComputerName           = $Env:ComputerName
        ISO8601                = (Get-Date $CurrTime -Format "yyyy-MM-dd HH:mm:ss")
        LocalDate              = $CurrTime.ToLongDateString()
        UTCDate                = $CurrTime.ToUniversalTime().ToLongDateString()
        LocalTime              = $CurrTime.ToLongTimeString()
        UTCTime                =  $CurrTime.ToUniversalTime().ToLongTimeString()
        TimeZone               = $ZoneInfo.Id
        DisplayName            = $ZoneInfo.DisplayName
        StandardName           = $ZoneInfo.StandardName
        DaylightName           = $ZoneInfo.DaylightName
        UTCOffset              = $ZoneInfo.BaseUTCOffset
        SupportsDaylightSaving = $ZoneInfo.SupportsDaylightSavingTime
        TimeServer             = (($NTP | Select-String "Source:").ToString()).Replace("Source:",$Null).Trim()
    })
    
}
} #end Get-PKWindowsDateTime