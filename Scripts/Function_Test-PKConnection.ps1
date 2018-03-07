#requires -version 3
function Test-PKConnection {
<#
.SYNOPSIS
    Uses Get-WMIObject and the Win32_PingStatus class to quickly ping one or more targets, either interactively or as a PSJob

.DESCRIPTION
    Uses Get-WMIObject and the Win32_PingStatus class to quickly ping one or more targets, either interactively or as a PSJob
    Accepts pipeline input
    Outputs a PSObject or PSJob

.Notes
    Name    : Function_Test-PKConnection.ps1
    Created : 2018-02-20
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
        
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2018-02-20 - Created script based on original in Idera link

.LINK
    http://community.idera.com/powershell/powertips/b/tips/posts/creating-highspeed-ping-part-4

.PARAMETER ComputerName
    One or more computer names or target IPs

.PARAMETER TimeoutMillisec
    Timeout in milliseconds (default is 1000)

.PARAMETER AsJob
    Run as local PSJob using Start-Job

.PARAMETER SuppressConsoleOutput
    Suppress all non-verbose/non-error console output

.EXAMPLE
    Test-PKConnection -ComputerName ops-vmtest-1 -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value            
        ---                   -----            
        ComputerName          {ops-vmtest-1}   
        Verbose               True             
        TimeoutMillisec       1000             
        AsJob                 False            
        SuppressConsoleOutput False            
        PipelineInput         False            
        ScriptName            Test-PKConnection
        ScriptVersion         1.0.0            

        VERBOSE: Test fast ping via WMI using -millisecond timeout
        VERBOSE: Pinging 1 target(s)

        Address       Alive  StatusCode
        -------       -----  ----------
        ops-vmtest-1  True            0

.EXAMPLE
    PS C:\> $Arr | Test-PKConnection -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value            
        ---                   -----            
        Verbose               True             
        ComputerName                           
        TimeoutMillisec       1000             
        AsJob                 False            
        SuppressConsoleOutput False            
        PipelineInput         True             
        ScriptName            Test-PKConnection
        ScriptVersion         1.0.0            

        VERBOSE: Test fast ping via WMI using 1000-millisecond timeout

        Address                      Alive StatusCode
        -------                      ----- ----------
        server14.domain.local        True          0
        webserver11.domain.local     True          0
        workstation11.domain.local   True          0
        10.11.12.13                  True          0
        legacyweb.domain.local       False     11010
        sql-dev.domain.local         True          0
        google.com                   True          0

#>
[cmdletbinding()]
param(
    [Parameter(
        ValueFromPipeline,
        Mandatory,
        Position = 0,
        HelpMessage = "One or more computer names or IP addresses"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerName,
    
    [Parameter(
        Position = 1,
        HelpMessage = "Timeout in milliseconds for ping (default is 1000)"
    )]
    [ValidateNotNullOrEmpty()]
    [int]$TimeoutMillisec = 1000,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory = $False,
        HelpMessage = "Run as PSJob"
    )]
    [Switch] $AsJob,

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
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
        
    # Inner function if not running as job
    If (-not $AsJob.IsPresent) {
        Function GetStatus{
            If ($_.StatusCode -eq 0){$True}
            Else {$False}
        }    
    }

    $Msg = "Test fast ping via WMI using $TimeoutMillisec-millisecond timeout"
    If ($AsJob.IsPresent) {$Msg += " as PowerShell job"}
    $Activity = $Msg
    Write-Verbose $Msg
}
Process {    
    
    # convert list of computers into a WMI query string
    If ($PipelineInput) {
        $query = $_ -join "' or Address='"
        $Msg = "Query is $Query"
        Write-Verbose $Msg
        
    }
    Else {
        $query = $ComputerName -join "' or Address='"
        $Msg = "Pinging $($ComputerName.Count) target(s)"
        Write-Verbose $Msg
        $Msg = "Query is $Query"
        Write-Verbose $Msg
    }

    Write-Progress -Activity $Activity -CurrentOperation $Msg

    If ($AsJob.IsPresent) {
        $SB = {
            Param($Query,$TimeoutMillisec)
                Function GetStatus{
                    If ($_.StatusCode -eq 0){$True}
                    Else {$False}
                }
            Try {
                Get-WmiObject -Class Win32_PingStatus -Filter "(Address='$Using:query') and timeout=$Using:TimeoutMillisec" -EA SilentlyContinue | 
                    Select-Object -Property Address,@{N="Alive";E={GetStatus}},StatusCode
            }
            Catch {
                $Host.UI.WriteErrorLine($_.Exception)
            }
        }
        Start-Job -ScriptBlock $SB
        #Invoke-Command  -AsJob
    }
    Else {
        Try {
            Get-WmiObject -Class Win32_PingStatus -Filter "(Address='$query') and timeout=$TimeoutMillisec" -EA SilentlyContinue | 
                Select-Object -Property Address,@{N="Alive";E={GetStatus}},StatusCode
        }
        Catch {
            $Host.UI.WriteErrorLine($_.Exception)
        }
    }
}

End { 
    Write-Progress -Activity $Activity -CurrentOperation $Msg -Completed
}
} #end Invoke-PKFastPing
