#requires -Version 3
function ConvertFrom-ErrorRecord {
<#
.SYNOPSIS
    Converts an error record or stop exception to a readable format.

.DESCRIPTION
    Converts an error record or stop exception to a readable format.
    Whenever PowerShell encounters an error, it emits an Error Record with detailed information about the problem. 
    Unfortunately, these objects are a bit cryptic and won’t show all of their information by default; this function makes
    them more intelligible.
    Supports pipeline input

.OUTPUT
    PSObject

.LINK
    https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/making-error-records-more-readable

.EXAMPLE
    PS C:\> $Error | ConvertFrom-ErrorRecord | Out-GridView

#>

  param(

    [Parameter(
        Mandatory,ValueFromPipeline,
        ParameterSetName='ErrorRecord',
        HelpMessage = "Error record"
    )]
    [Management.Automation.ErrorRecord[]]$ErrorRecord,

    [Parameter(
        Mandatory,ValueFromPipeline,
        ParameterSetName='StopException',
        HelpMessage = "A special stop exception raised by cmdlets with -ErrorAction Stop"
    )]
    [Management.Automation.ActionPreferenceStopException[]]$Exception

  )


  process {

    # If we received a stop exception in $Exception, the error record is to be found inside of it
    # In all other cases, $ErrorRecord was received directly

    If ($PSCmdlet.ParameterSetName -eq 'StopException'){
        $ErrorRecord = $Exception.ErrorRecord
    } 

    # Compose a new object out of the interesting properties found in the error record object
    $ErrorRecord | ForEach-Object { 
        [PSCustomObject]@{
        Exception = $_.Exception.Message
        Reason    = $_.CategoryInfo.Reason
        Target    = $_.CategoryInfo.TargetName
        Script    = $_.InvocationInfo.ScriptName
        Line      = $_.InvocationInfo.ScriptLineNumber
        Column    = $_.InvocationInfo.OffsetInLine
      }
    }
}

} # end  