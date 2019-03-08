#requires -Version 3
function ConvertFrom-PKErrorRecord {
<#
.SYNOPSIS
    Converts an error record or stop exception to a more intelligible format

.DESCRIPTION
    Converts an error record or stop exception to a more intelligible format
    Supports pipeline input

.NOTES
    Name    : Function_ConvertFrom-PKErrorRecord.ps1 
    Created : 2019-03-08
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.01.0000 - 2019-03-08 - Created from Idera original
     
.LINK
    https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/making-error-records-more-readable

.EXAMPLE
    PS C:\> ConvertFrom-PKErrorRecord -ErrorRecord $Error[0]

        VERBOSE: PSBoundParameters: 
	
        Key           Value                                                       
        ---           -----                                                       
        ErrorRecord   {fatal: pathspec 'Drafts/Hashtable' did not match any files}
        Exception                                                                 
        ScriptName    ConvertFrom-PKErrorRecord                                   
        ScriptVersion 1.0.0                                                       

        Exception : fatal: pathspec 'Drafts/Hashtable' did not match any files
        Reason    : WriteErrorException
        Target    : 
        Script    : 
        Line      : 1
        Column    : 1

.EXAMPLE
    PS C:\> $Error | ConvertFrom-PKErrorRecord

        Exception : An error occurred while enumerating through a collection: Collection was modified; enumeration operation may not execute..
        Reason    : RuntimeException
        Target    : System.Collections.ArrayList+ArrayListEnumeratorSimple
        Script    : 
        Line      : 1
        Column    : 1

        Exception : The remote server returned an error: (401) Unauthorized.
        Reason    : WebException
        Target    : System.Net.HttpWebRequest
        Script    : C:\Users\jbloggs\repos\dellwarranty.ps1
        Line      : 741
        Column    : 42

        Exception : Parameter set cannot be resolved using the specified named parameters.
        Reason    : ParameterBindingException
        Target    : System.Management.Automation.ParseException: At line:1 char:2
                    +  | fl
                    +  ~
                    An empty pipe element is not allowed.
                       at System.Management.Automation.Runspaces.PipelineBase.Invoke(IEnumerable input)
                       at System.Management.Automation.PowerShell.Worker.ConstructPipelineAndDoWork(Runspace rs, Boolean performSyncInvoke)
                       at System.Management.Automation.PowerShell.CoreInvokeHelper[TInput,TOutput](PSDataCollection`1 input, PSDataCollection`1 output, 
                    PSInvocationSettings settings)
                       at System.Management.Automation.PowerShell.CoreInvoke[TInput,TOutput](PSDataCollection`1 input, PSDataCollection`1 output, 
                    PSInvocationSettings settings)
                       at System.Management.Automation.PowerShell.InvokeWithDebugger(IEnumerable`1 input, IList`1 output, PSInvocationSettings settings, 
                    Boolean invokeMustRun)
                       at System.Management.Automation.ScriptDebugger.ProcessCommand(PSCommand command, PSDataCollection`1 output)
                       at Microsoft.PowerShell.Host.ISE.PowerShellTab.PowerShellDebugInvoker.<>c__DisplayClass19_0.<BeginInvoke>b__0()
                       at Microsoft.PowerShell.Host.ISE.PowerShellTab.EnterNestedPrompt()
        Script    : 
        Line      : 1
        Column    : 10


#>
[CmdletBinding()]
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
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Display our parameters
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

}
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

$Null = New-Alias -Name ConvertFrom-ErrorRecord -Value ConvertFrom-PKErrorRecord -Force -Verbose:$False -Confirm:$False