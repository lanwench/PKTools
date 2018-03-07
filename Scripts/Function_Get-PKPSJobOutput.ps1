#requires -Version 3
Function Get-PKCompletedJobOutput{
<#
.SYNOPSIS
    Gets the results for completed PowerShell jobs by name or ID, with option to keep or remove results, or remove job entirely

.DESCRIPTION
    Gets the results for completed PowerShell jobs by name or ID, with option to keep or remove results, or remove job entirely
    SupportsShouldProcess
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES
    Name    : Function_Get-PKCompletedPSJobOutput.ps1
    Created : 2018-01-27
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2018-01-27 - Created script

.PARAMETER JobName
    One or more job names(default is all completed jobs)

.PARAMETER JobName
    One or more job IDs

.PARAMETER JobResultAction
    After getting job results, either remove or keep them (default is keep)

.PARAMETER RemoveJob
    Remove job (incompatible with KeepResults)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKCompletedJobOutput -Verbose
    # Get results of all completed jobs, keep results, keep jobs

        VERBOSE: PSBoundParameters:

        Key                   Value
        ---                   -----
        Verbose               True
        JobName
        JobID
        JobResultAction       KeepResults
        RemoveJob             False
        SuppressConsoleOutput False
        PipelineInput         False
        ParameterSetName      Name
        ScriptName            Get-PKCompletedJobOutput
        ScriptVersion         1.0.0

        Action: Get PSJob output (keep job results)
        VERBOSE: Get all completed jobs
        VERBOSE: 4 matching completed job(s) found
        No result data available for job ID 26: name Job26, type RemoteJob, location webserver13
        VERBOSE: Job ID 28: name GWMI, type RemoteJob, location sql2012
        VERBOSE: Job ID 30: name gwmi, type RemoteJob, location sql2014,webserver
        VERBOSE: Job ID 33: name Uptime, type RemoteJob, location testvm


        SystemDirectory : C:\Windows\system32
        Organization    : Amalgamated Toothpick, Inc.
        BuildNumber     : 9600
        RegisteredUser  : Operations
        SerialNumber    : 00253-40420-45537-AA519
        Version         : 6.3.9600
        PSComputerName  : sql2012

        SystemDirectory : C:\Windows\system32
        Organization    : Amalgamated Toothpick, Inc.
        BuildNumber     : 9600
        RegisteredUser  : Operations
        SerialNumber    : 00253-40420-45537-AA130
        Version         : 6.3.9600
        PSComputerName  : sql2014

        SystemDirectory : C:\Windows\system32
        Organization    : Microsoft
        BuildNumber     : 7601
        RegisteredUser  : AutoBVT
        SerialNumber    : 00477-OEM-8400101-10502
        Version         : 6.1.7601
        PSComputerName  : webserver

        PSComputerName    : testvm
        RunspaceId        : 2640c71e-848d-4dcb-950e-3ff2d21c421f
        Ticks             : 33012096866300
        Days              : 38
        Hours             : 5
        Milliseconds      : 686
        Minutes           : 0
        Seconds           : 9
        TotalDays         : 38.2084454471065
        TotalHours        : 917.002690730556
        TotalMilliseconds : 3301209686.63
        TotalMinutes      : 55020.1614438333
        TotalSeconds      : 3301209.68663

.EXAMPLE
    PS C:\> Get-PKCompletedJobOutput -JobName gw* -RemoveJob

        WARNING: -RemoveJob cannot be used when JobResultAction is set to 'KeepResults'; jobs will not be removed
        Action: Get PSJob output (keep job results)

        SystemDirectory : C:\Windows\system32
        Organization    : Amalgamated Toothpick
        BuildNumber     : 9600
        RegisteredUser  : Operations
        SerialNumber    : 00253-40420-45537-AA519
        Version         : 6.3.9600
        PSComputerName  : tempvm

        SystemDirectory : C:\Windows\system32
        Organization    : Amalgamated Toothpick
        BuildNumber     : 9600
        RegisteredUser  : Operations
        SerialNumber    : 00253-40420-45537-AA130
        Version         : 6.3.9600
        PSComputerName  : sqldev

.EXAMPLE
    PS C:\> Get-PKCompletedJobOutput -JobID 26 -Verbose

        VERBOSE: PSBoundParameters:

        Key                   Value
        ---                   -----
        JobID                 {26}
        Verbose               True
        JobName
        JobResultAction       KeepResults
        RemoveJob             False
        SuppressConsoleOutput False
        PipelineInput         False
        ParameterSetName      ID
        ScriptName            Get-PKCompletedJobOutput
        ScriptVersion         1.0.0

        Action: Get PSJob output (keep job results)
        VERBOSE: Get completed jobs by ID
        VERBOSE: 1 matching completed job(s) found
        VERBOSE: Job ID 26 (Job26)
        No result data available for job ID 26 (Job26)
        No job output returned

.EXAMPLE
    PS C:\> Get-PKCompletedJobOutput -JobID 26 -JobResultAction RemoveResults -RemoveJob -Verbose

        VERBOSE: PSBoundParameters:

        Key                   Value
        ---                   -----
        JobID                 {26}
        JobResultAction       RemoveResults
        RemoveJob             True
        Verbose               True
        JobName
        SuppressConsoleOutput False
        PipelineInput         False
        ParameterSetName      ID
        ScriptName            Get-PKCompletedJobOutput
        ScriptVersion         1.0.0

        Action: Get PSJob output (remove job object)
        VERBOSE: Get completed jobs by ID
        VERBOSE: 1 matching completed job(s) found
        ERROR: No result data available for job ID 26: name Job26, type RemoteJob, location remote-dev
        VERBOSE: Remove job
        VERBOSE: Removed job ID 26

#>
[CmdletBinding(
    DefaultParameterSetName = "Name",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        ParameterSetName = "Name",
        Position = 0,
        Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Job name (default is all)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name")]
    [String[]] $JobName,

    [Parameter(
        ParameterSetName = "ID",
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Job ID - at least one is mandatory"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("ID")]
    [String[]] $JobID,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Keep or remove PSJob results after Receive-Job (default is keep)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("KeepResults","RemoveResults")]
    [string]$JobResultAction = "KeepResults",

    [Parameter(
        Mandatory=$False,
        HelpMessage="Remove PSJob object after returning results (incompatible with KeepJobResults)"
    )]
    [switch]$RemoveJob,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Detect pipeline input
    $PipelineInput = $False
    $KeyNames = "JobName","JobID"
    $Pipeline = $KeyNames | Foreach-Object {
        ((-not $PSBoundParameters.ContainsKey($_)) -and (-not $_))
    }
    If ($Pipeline -contains $True) {$PipelineInput = $True}

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} |
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    If ($JobResultAction -eq "KeepResults") {
        If ($RemoveJob.IsPresent) {
            $Msg = "-RemoveJob cannot be used when JobResultAction is set to 'KeepResults'; jobs will not be removed"
            Write-Warning $Msg
            $RemoveJob = $False
        }
    }
    Else {
        # If default for Receive-Job is set to -Keep, temporarily remove it
        If ($PSDefaultParameterValues.Keys -contains "Receive-Job:Keep"){
            [switch]$ChangedPref = $True
            $PSDefaultParameterValues.Remove("Receive-Job:Keep")
        }
    }

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Get-Job using name (or all names)
    $Param_GetJobName = @{}
    $Param_GetJobName = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }
    If ($JobName) {$Param_GetJobName.Add("Name",$Null)}

    # Splat for Get-Job using ID
    $Param_GetJobID = @{}
    $Param_GetJobID = @{
        ID          = $Null
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Receive-Job
    $Param_Receive = @{}
    $Param_Receive = @{
        Keep        = $False
        ErrorAction = "Stop"
        Verbose     = $False
    }
    If ($JobResultAction -eq "KeepResults") {
        $Param_Receive.Add("WarningAction","SilentlyContinue")
        $Param_Receive.Keep = $True
    }

    # Splat for Remove-Job
    If ($RemoveJob.IsPresent) {
        $Param_Remove = @{}
        $Param_Remove = @{
            Force       = $True
            Confirm     = $False
            ErrorAction = "Stop"
            Verbose     = $False
        }
    }


    #Splat for Write-Progress
    $Activity = "Get PSJob output"
    If ($RemoveJob.IsPresent) {$Activity += " (remove job object)"}
    ElseIf ($JobResultAction -eq "KeepResults") {$Activity += " (keep job results)"}
    $Param_WP = @{
        Activity         = $Activity
        Status           = "Working"
        CurrentOperation = $Null
    }

    #endregion Splats

    #Output
    [array]$JobResults = @()

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Action = $Activity
    $Msg = "Action: $Action"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}

Process {

    Switch ($Source) {
        Name {
            Try {
                If ($JobName) {
                    $Msg = "Get completed jobs by name"
                    $Param_GetJobName.Name = $JobName
                }
                Else {
                    $Msg = "Get all completed jobs"
                }

                $Param_WP.CurrentOperation = $Msg
                Write-Progress @Param_WP
                Write-Verbose $Msg

                [array]$Jobs = Get-Job @Param_GetJobName | Where-Object {$_.State -eq "Completed"}
                $Msg = "$($Jobs.Count) matching completed job(s) found"
                Write-Verbose $Msg
            }
            Catch {
                $Msg = "No matching completed jobs found"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        ID {
            Try {
                $Msg = "Get completed jobs by ID"
                $Param_WP.CurrentOperation = $Msg
                Write-Verbose $Msg
                Write-Progress @Param_WP

                $Param_GetJobID.ID = $JobID
                [array]$Jobs = Get-Job @Param_GetJobID | Where-Object {$_.State -eq "Completed"}

                $Msg = "$($Jobs.Count) matching completed job(s) found"
                Write-Verbose $Msg
            }
            Catch {
                $Msg = "No matching jobs found"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine($Msg)
            }
        }
    }

    If ($Jobs.Count -gt 0) {

        $Total = $Jobs.count
        $Current = 0

        $Status = "Get job results"

        Foreach ($Job in $Jobs) {

            $Current ++

            $Curr = "ID $($Job.ID): name $($Job.Name), type $($Job.PSJobTypeName), location $($Job.Location)"
            $Param_WP.Status = $Curr
            $CurrentOp = "Get job results"
            If ($KeepJobResults.IsPresent) {$CurrentOp += " (keep results)"}
            $Param_WP.CurrentOperation = $CurrentOp

            Write-Progress @Param_WP -PercentComplete ($Current/$Total*100)

            Try {

                If ($Job.HasMoreData -eq $True) {
                    Write-Verbose "Job $Curr"
                    $JobResults += $Job | Receive-Job @Param_Receive
                }
                Else {
                    $Msg = "No result data available for job $Curr"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                }

                If ($RemoveJob.IsPresent) {

                    $Msg = "Remove job"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP -PercentComplete ($Current/$Total*100)
                    Write-Verbose $Msg

                    $ConfirmMsg = "`n`n`tRemove job $Curr`n`n"
                    If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$ConfirmMsg)) {
                        Try {
                            $Job | Remove-Job @Param_Remove
                            $Msg = "Removed job $($JobID)"
                            Write-Verbose $Msg
                        }
                        Catch {
                            $Msg = "Failed to remove job $($JobID)"
                            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                            $Host.UI.WriteErrorLine($Msg)
                        }
                    }
                    Else {
                        $Msg = "Job removal cancelled by user for job $($Job.ID)"
                        $Host.UI.WriteErrorLine($Msg)
                    }
                } # end if removing job
            }
            Catch {
                $Msg = "Failed to get results  for job $($Job.ID)"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine($Msg)
            }

        } #end foreach

    } # end if jobs

}
End {

    Write-Progress -Activity $Activity -Completed

    # Reset the default job action preference
    If ($ChangedPref.IsPresent) {
        If (-not ($PSDefaultParameterValues.Keys -contains "Receive-Job:Keep")){
            $PSDefaultParameterValues.Add("Receive-Job:Keep",$True)
        }
        Elseif ($PSDefaultParameterValues.Keys."Receive-Job:Keep" -ne $True) {
            $PSDefaultParameterValues."Receive-Job:Keep" = $True
        }
    }

    If ($JobResults.Count -gt 0) {
        Write-Output ($JobResults | Select -Property * -ExcludeProperty PSComputerName,RunspaceID)
    }

}

} #end Get-PKCompletedPSJobOutput

$Null = New-Alias -Name "Get-PKJobOutput" -Value Get-PKCompletedPSJobOutput -Description "For simplicity 2018-01-27" -Force -ErrorAction SilentlyContinue -Verbose:$False -Confirm:$False