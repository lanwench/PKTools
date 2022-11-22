#requires -Version 3
Function Invoke-Something {
<#
.SYNOPSIS 
    Template for running a scriptblock on a local or remote computer, interactively or as a PSJob

.DESCRIPTION
    Template for running a scriptblock on a local or remote computer, interactively or as a PSJob
    Runs THING in a scriptblock
    Accepts pipeline input
    Returns a PSObject or PSJob

.NOTES
    Name      : Function_Invoke-Something.ps1
    Created   : 2022-08-12
    Author    : Paula Kingsley
    Version   : 01.00.0000
    Changelog :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2022-08-12 - Created script

.PARAMETER ComputerName
    One or more Windows computer names (default is local computer)

.PARAMETER Detailed
    Return additional properties in output

.PARAMETER Credential
    Valid admin credentials on computer

.PARAMETER Authentication
    Authentication protocol: Basic, CredSSP, Negotiate, Kerberos (default is Negotiate)

.PARAMETER ConnectionTest
    Connectivity test options prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests are ignored on local computer)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PowerShell job instead of interactively

.EXAMPLE
    PS C:\> 
#>
[Cmdletbinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more Windows computer names (default is local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name","DNSHostName")]
    [object[]]$ComputerName,

    [Parameter(
        HelpMessage = "Return additional properties in output"
    )]
    [switch]$Detailed,

    [Parameter(
        HelpMessage = "Valid admin credentials on computer"
    )]
    [ValidateNotNullOrEmpty()]
    [PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "Authentication protocol for WinRM: Basic, CredSSP, Negotiate, Kerberos (default is Negotiate)"
    )]
    [ValidateSet("Basic","CredSSP","Negotiate","Kerberos")]
    [string]$Authentication = "Negotiate",

    [Parameter(
        HelpMessage = "Connectivity test options prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests are ignored on local computer)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",
    
    [Parameter(
        HelpMessage = "Run Invoke-Command scriptblock as PowerShell job instead of interactively"
    )]
    [switch]$AsJob
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    If (-not $CurrentParams.Computername) {$ComputerName = $CurrentParams.ComputerName = $Env:ComputerName}
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Connectivity test using Test-WSMAN
    Function Test-WinRM{
        Param($Computer)
        $Param_WSMAN = @{
            ComputerName   = $Computer
            Credential     = $Credential
            Authentication = $Authentication
            ErrorAction    = "Silentlycontinue"
            Verbose        = $False
        }
        Try {
            If (Test-WSMan @Param_WSMAN) {$True}
            Else {$False}
        }
        Catch {$False}
    }
    
    # Connectivity test using ping
    Function Test-Ping{
        Param($Computer)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
    }

    # Scriptblock to get activation status via WMI; using GWMI for backwards compatibility
    $Scriptblock = {
        Param($Detailed)
        
        # Something happens here
    }
    
    # Splats
    $Param_IC = @{
        Credential     = $Credential
        Authentication = $Authentication
        ArgumentList   = $Detailed
        Scriptblock    = $ScriptBlock
        AsJob          = $AsJob
        ErrorAction    = "Stop"
    }

    $Activity = "Do a thing"
    $CurrentOperation = "Invoke scriptblock"
    If ($AsJob.IsPresent) {
        $CurrentOperation += " as PowerShell job"
        $Jobs = @()
        $JobPrefix = "thing"
    }
    Write-Verbose "[BEGIN: $ScriptName] $Activity"
}
Process {
    
    $Total = $Computername.Count
    $Current = 0
    Foreach ($Computer in $ComputerName) {
        
        $Current ++ 
        If ($Computer -is [string]) {
            $Computer = $Computer.Trim()
        }
        Elseif ($Computer -is [Microsoft.ActiveDirectory.Management.ADAccount]) {
            If ($Computer.DNSHostName) {
                $Computer = $Computer.DNSHostName
            }
            Else {
                $Computer = $Computer.Name
            }
        }

        # Flag
        [switch]$Continue = $False
        [switch]$IsLocal = $False
        
        # Test connectivity if not local
        If ($Computer -match ("^$($Env:ComputerName)$|^$($Env:ComputerName)\.|localhost|127.0.0.1")) {
            $Continue = $True
            $IsLocal = $True
        }
        Else {
            Switch ($ConnectionTest) {
                Default {$Continue = $True}
                Ping {
                    $Msg = "Ping computer"
                    Write-Verbose "[$Computer] $Msg" 
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                    If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                    Else {
                        $Msg = "Ping failure"
                        Write-Warning "[$Computer] $Msg"
                    }
                }
                WinRM {
                    $Msg = "Test WinRM connection"
                    Write-Verbose "[$Computer] $Msg" 
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                    If ($Null = Test-WinRM -Computer $Computer) {
                        $Continue = $True
                    }
                    Else {
                        $Msg = "WinRM failure"
                        Write-Warning "[$Computer] $Msg" 
                    }
                }        
            } #end switch
        }

        If ($Continue.IsPresent) {
            $Msg = "Invoke scriptblock"
            If ($AsJob.IsPresent) {$Msg += " (as Powershell job)"}
            Write-Verbose "[$Computer] $Msg" 
            Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
            Try {
                If ($Computer -eq $Env:Computername) {
                    If ($AsJob.IsPresent) {
                        $Jobs += Start-Job -ScriptBlock $Scriptblock -ArgumentList $Detailed -Name "$JobPrefix_$Computer" -ErrorAction Stop
                    }
                    Else {
                        Invoke-Command -ScriptBlock $Scriptblock -ArgumentList $Detailed -ErrorAction Stop | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID
                    }
                }
                Else {
                    If ($AsJob.IsPresent) {
                        $Jobs += Invoke-Command -Computername $Computer @Param_IC -JobName "$JobPrefix_$Computer"
                    }
                    Else {
                        Invoke-Command -Computername $Computer @Param_IC | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID
                    }
                }
            }
            Catch {
                $Msg = "Operation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                Write-Warning "[$Computer] $Msg"
            }
        
    } #end if continue

    } # end foreach

}
End {
    If ($AsJob.IsPresent) {
        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; try running 'Get-Job -Id # | Wait-Job | Receive-Job | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID' to view output`n"
            Write-Verbose $Msg
            $Jobs | Get-Job
        }
        Else {
            $Msg = "No jobs created"
            Write-Warning $Msg
        }
    }

    $Null = Write-Progress -Activity * -Completed
    $Msg = "[END: $Scriptname] $Activity" 
    Write-Verbose $Msg
}
} #end Invoke-Something