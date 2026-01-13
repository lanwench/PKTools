#reqiures -Version 4
Function Get-PKGoogleFSLogErrors {
<#
.SYNOPSIS
    Scans local computer Google Drive FileSync log files for errors and returns matching entries.

.DESCRIPTION
    The Get-PKGoogleFSLogErrors function searches one or more Google Drive FileSync (DriveFS) log files for entries containing error patterns indicating failures, fatal errors, crashes, or dumps. 
    By default it excludes matches that clutter the view and can provide misleading results.
    The GroupByLine parameter will group the output based on the text matched line content and sort by occurrence count, useful for identifying recurring issues.

.NOTES
    Name    : Function_Get-PKGoogleFSLogErrors.ps1
    Created : 2025-12-22
    Author  : Paula Kingsley
    Version : 01.00.1000
    History: 
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2025-12-22 - Created script


.PARAMETER LogPath
    Specifies the path to the Google Drive FileSync logs directory (default is the user's AppData Local Google DriveFS Logs folder).

.PARAMETER Pattern
    Specifies the regex pattern to search for in log files (default is 'fail|fatal|error|crash|dump')

.PARAMETER Exclude
    Specifies a pattern for log entries to exclude from results (default is 'AUTOMATIC_ERROR_REPORTING_ENABLED')

.PARAMETER NumFiles
    Specifies the number of most recent log files to scan (default is 1)

.PARAMETER GroupByLine
    When specified, groups duplicate error messages and sorts by occurrence count instead of returning individual lines.

.OUTPUTS
    PSCustomObject
    Returns custom objects with FileName, Date, LineNumber, and Line properties (if -GroupByLine, Instances, String, and FileName)

.EXAMPLE
    PS C:\> Get-PKGoogleFSLogErrors
    Scans the most recent Google Drive log file for default error patterns.

.EXAMPLE
    PS C:\> Get-PKGoogleFSLogErrors -NumFiles 5 -GroupByLine
    Scans the 5 most recent log files, groups duplicate errors, and displays them sorted by frequency.

.EXAMPLE
    PS C:\> Get-PKGoogleFSLogErrors -Pattern "sync.*failed" -Exclude "warning" -NumFiles 3
    Searches 3 recent log files for "sync.*failed" pattern while excluding lines containing "warning".


#>

    [CmdletBinding()]
    Param(
        [Parameter(
            Position = 0,
            HelpMessage = "Path to Google Drive FileSync log files"
        )]
        [string]$LogPath = "$Env:UserProfile\AppData\Local\Google\DriveFS\Logs",

        [Parameter(
            HelpMessage = "Case-insensitive regex pattern to search for in log files (default is 'fail|fatal|error|crash|dump')"
        )]
        [string]$Pattern = 'fail|fatal|error|crash|dump',
        
        [Parameter(
            HelpMessage = "Case-insensitive regex pattern to exclude from results (default is 'AUTOMATIC_ERROR_REPORTING_ENABLED')"
        )]
        [string]$Exclude = 'AUTOMATIC_ERROR_REPORTING_ENABLED',

        [Parameter(
            HelpMessage = "Number of most recent log files to scan (default is 1)"
        )]
        [int]$NumFiles = 1,

        [Parameter(
            HelpMessage = "Group duplicate error messages and sort by occurrence count"
        )]
        [switch]$GroupByLine
    )
    Begin {
        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # Show our settings
        $ScriptName = $MyInvocation.MyCommand.Name
        $CurrentParams = $PSBoundParameters
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("PipelineInput", $PipelineInput)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        Write-Verbose "[BEGIN: $ScriptName] Get certificate details on port $Port"
    }
    Process {
        Write-Verbose "Scanning $LogPath for last $NumFiles Google Drive client log files containing errors"
        Try {
            $Logs = Get-Item $LogPath -ErrorAction Stop
            $LatestLogs = $Logs | Get-ChildItem -Filter "drive_fs*.txt" -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First $NumFiles
            $Results = Get-Item $LatestLogs -PipelineVariable OrigFile | 
                Select-String -Pattern $Pattern -Exclude $Exclude | Select-Object @{N="Name";E={$OrigFile.Name}},
                    @{N="Date";E={Get-Date $OrigFile.LastWriteTime -Format 'yyyy-MM-dd HH:mm:ss'}}, 
                    LineNumber, 
                    @{N="Line";E={$_.Line.Trim()}} | Sort-Object Date -Descending  # ,@{N="FullName";E={$OrigFile.FullName}}
            If ($GroupByLine.IsPresent) {
                Write-Verbose "Grouping results by duplicated line content, sorted by count"
                $Results | Group-Object Line | Where-Object {$_.Count -gt 1} | Sort-Object Count -Descending | Select-Object @{N="Instances";E={$_.Count}},@{N="String";E={$_.Name}},@{N="FileName";E={(@($_.Group.FileName | Select-Object -Unique).TrimStart()) -join(",")}}
            }
            Else {Write-Output $Results}
        }
        Catch {
            Write-Warning $_.Exception.Message
        }
    }
} # end Get-PKGoogleFSLogErrors

