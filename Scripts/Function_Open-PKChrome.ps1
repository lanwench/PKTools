#requires -Version 4
Function Open-PKChrome{
<#
.SYNOPSIS
    Launches a URL in Chrome, with options for default profile/new window 
    
.DESCRIPTION
    Launches a URL in Chrome, with options for default profile/new window
    By default opens a new tab in the current window if open
    
.NOTES        
    Name    : Function_Open-PKChrome.ps1
    Created : 2023-10-25
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
            
        v01.00.0000 - 2023-10-25 - Created script
        
.EXAMPLE
    PS C:\> Open-PKChromeWindow -URL https://github.com/lanwench/PKTools.git -Verbose -NewWindow
    
#>
[cmdletbinding()]
Param(
    [Parameter(
        Mandatory
    )]
    [ValidateNotNullOrEmpty()]
    [string]$URL,
    [switch]$UseDefaultProfile,
    [switch]$NewWindow
)
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"
    
    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
    Where-Object { Test-Path variable:$_ } | Foreach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    Write-Verbose "[BEGIN: $ScriptName]"
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
}