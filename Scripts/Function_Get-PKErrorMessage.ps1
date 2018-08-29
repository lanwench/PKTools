#requires -Version 3
<#
.SUMMARY
    Returns details about errors from ErrorRecord objects

.DESCRIPTION
    Returns details about errors from ErrorRecord objects
    Accepts pipeline input
    Outputs a PSObject

.NOTES
    Name    : Function_Get-PKErrorMessage
    Created : 2018-05-17
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2018-05-17 - Created script based on Idera PowerTip
        v01.01.0000 - 2018-08-29 - Removed test line at bottom

.LINK
    http://community.idera.com/powershell/powertips/b/tips/posts/converting-error-records

.EXAMPLE
    PS C:\> $Error[22] | Get-PKErrorMessage 

        Exception  : Cannot find a variable with the name 'SourceLocation'.
        Reason     : ItemNotFoundException
        Target     : SourceLocation
        Command    : Get-Variable
        Script     : C:\Users\pkingsley\Dropbox\Powershell\Modules\PowerShellGet\PowerShellGet\PSModule.psm1
        LineNumber : 4299
        Column     : 20
        Line       :                 if(Get-Variable -Name SourceLocation -ErrorAction SilentlyContinue)

#>

Function Get-PKErrorMessage {
  [CmdletBinding(DefaultParameterSetName="ErrorRecord")]
  param(
    
    [Parameter(
        ValueFromPipeline,
        ParameterSetName = "ErrorRecord",
        Position = 0
    )]
    [ValidateNotNullOrEmpty()]
    [Management.Automation.ErrorRecord]$Record = $Error[0],
    
    [Parameter(
        Mandatory,
        ValueFromPipeline,
        ParameterSetName = "Unknown", 
        Position = 0
    )]
    [Object]$Alien
  )
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Msg = "Return error details"
    Write-Verbose $Msg
}
  
Process{

    Switch ($Source) {
        ErrorRecord {
            [PSCustomObject]@{
                Exception  = $Record.Exception.Message
                Reason     = $Record.CategoryInfo.Reason
                Target     = $Record.CategoryInfo.TargetName
                Command    = $Record.InvocationInfo.MyCommand.Name
                Script     = $Record.InvocationInfo.ScriptName
                LineNumber = $Record.InvocationInfo.ScriptLineNumber
                Column     = $Record.InvocationInfo.OffsetInLine
                Line       = $Record.InvocationInfo.Line

            }
        }
        Default {
            Write-Warning "$Alien"
        }
    
    }
}
} # End Get-PKErrorMessage
