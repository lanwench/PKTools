#requires -Version 4
Function Open-PKChrome{
<#
.SYNOPSIS
    Launches a URL in Google Chrome with options for new window and default profile

.DESCRIPTION
    Launches Google Chrome and opens the specified URL
    By default opens in the current window if Chrome is already running
    Supports opening in a new window and overriding the profile directory


.NOTES        
    Name    : Function_Open-PKChrome.ps1
    Created : 2023-10-25
    Author  : Paula Kingsley
    Version : 01.01
    History :

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00 - 2023-10-25 - Created script
        v01.01 - 2026-06-17 - Cosmetic updates; converted version format to major.minor
        
.EXAMPLE
    PS C:\> Open-PKChrome -URL https://github.com/lanwench/PKTools.git -Verbose -NewWindow
    
#>
[CmdletBinding()]
Param(
    [Parameter(
        Mandatory,
        HelpMessage = "URL to open"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$URL,

    [switch]$UseDefaultProfile,
    [switch]$NewWindow
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01"
    
    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } |
    Where-Object { Test-Path variable:$_ } | ForEach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    Write-Verbose "[BEGIN: $ScriptName] Launch URL in Chrome"
}
Process {

    [string[]]$Arguments = $URL
    $Msg = "Launching $URL in Chrome"

    If ($NewWindow.IsPresent) {
        $Msg += " in new window"
        $Arguments += '--new-window'
    }
    If ($UseDefaultProfile.IsPresent) {
        $Msg += " using default profile"
        $Arguments += '--profile-directory="Default"'
    }
    Write-Verbose $Msg
    $Param = @{
        FilePath = "chrome.exe"
        ArgumentList = $Arguments
    }
    Try {
        Start-Process @Param
    }
    Catch {
        Throw $_.Exception.Message
    }
}
End {
    Write-Verbose "[END: $ScriptName] Script ran successfully"
}
}