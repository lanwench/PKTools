#Requires -version 3
Function Get-PKWindowsHardware {
<# 
.SYNOPSIS
    Does something cool, interactively or as a PSJob

.DESCRIPTION
    Does something cool, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Do-Somethingcool.ps1
    Created : 2019-03-08
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-10-22 - Created script

.PARAMETER ComputerName
    One or more computer names

.PARAMETER Credential
    Valid credentials on target (default is current user credentials)

.PARAMETER AsJob
    Invoke command as a PSjob

.PARAMETER ConnectionTest
    Run WinRM or ping connectivity test prior to Invoke-Command, or no test (default is WinRM)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKWindowsHardware -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Verbose               True                                     
        ComputerName          {WORKSTATION81}                        
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        JobPrefix             Hardware                                 
        ConnectionTest        WinRM                                    
        SuppressConsoleOutput False                                    
        ScriptName            Get-PKWindowsHardware                    
        ScriptVersion         1.0.0                                    

        BEGIN  : Invoke scriptblock to return Windows hardware details

        VERBOSE: [WORKSTATION81] Invoke command

        ComputerName : WORKSTATION81
        Manufacturer : Dell Inc.
        Model        : OptiPlex 9020
        SerialNumber : 4WIKW31
        Messages     : Operation completed successfully

        END    : Invoke scriptblock to return Windows hardware details

.EXAMPLE
    PS C:\> Get-VM ops-pk* | Get-PKWindowsHardware -Credential $Credential -AsJob -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Credential            System.Management.Automation.PSCredential
        AsJob                 True                                     
        ComputerName          {WORKSTATION81}                        
        JobPrefix             Hardware                                 
        ConnectionTest        WinRM                                    
        SuppressConsoleOutput False                                    
        ScriptName            Get-PKWindowsHardware                    
        ScriptVersion         1.0.0                                    

        BEGIN  : Invoke scriptblock to return Windows hardware details as remote PSJob

        VERBOSE: [ops-pktest-1] Test WinRM connection
        [ops-pktest-1] WinRM failure
        VERBOSE: [ops-pkbastion-1] Test WinRM connection
        VERBOSE: [ops-pkbastion-1] Invoke command as PSJob
        VERBOSE: [ops-pksql-1] Test WinRM connection
        VERBOSE: [ops-pksql-1] Invoke command as PSJob
        VERBOSE: 2019-03-08 11:49:46 AM	Get-VM	Finished execution
        VERBOSE: 2 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output
        
        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command  
        --     ----            -------------   -----         -----------     --------             -------  
        2      Hardware_ops... RemoteJob       Completed     True            ops-pkbastion-1      ...      
        4      Hardware_ops... RemoteJob       Completed     True            ops-pksql-1          ...    
          
        END    : Invoke scriptblock to return Windows hardware details as remote PSJob

        PS C:\> Get-Job | Receive-Job

        ComputerName   : OPS-PKBASTION-1
        Manufacturer   : No Enclosure
        Model          : VMware Virtual Platform
        SerialNumber   : None
        Messages       : Operation completed successfully
        PSComputerName : ops-pkbastion-1
        RunspaceId     : 411f248b-ea3a-4029-b4ba-bc5db2a3f739

        ComputerName   : OPS-PKSQL-1
        Manufacturer   : No Enclosure
        Model          : VMware Virtual Platform
        SerialNumber   : None
        Messages       : Operation completed successfully
        PSComputerName : ops-pksql-1
        RunspaceId     : bb3c889b-c0a9-4068-a19f-dad56eb93697

#> 

[CmdletBinding(
    DefaultParameterSetName = "__DefaultParameterSet",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName = $Env:ComputerName,

    [Parameter(
        HelpMessage="Valid credentials on target"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Run Invoke-Command scriptblock as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Prefix for job name (default is 'Hardware')"
    )]
    [String] $JobPrefix = "Hardware",

    [Parameter(
        HelpMessage="Connection test prior to Invoke-Command: WinRM using Kerberos), Ping, or None (default WinRM)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

    [Parameter(
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
  
    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        
        Try {
            $Param_GWMI = @{}
            $Param_GWMI = @{
                Class = $Null
                ErrorAction = "Stop"
            }
            
            $Param_GWMI.Class = "Win32_SystemEnclosure"
            $BIOS = Get-WMIObject @Param_GWMI
            $Param_GWMI.Class = "Win32_ComputerSystem"
            $System = Get-WMIObject @Param_GWMI
            
            New-Object PSObject -Property @{
                ComputerName = $Env:ComputerName
                Manufacturer = $BIOS.Manufacturer
                Model        = $System.Model
                SerialNumber = $BIOS.SerialNumber
                Messages     = "Operation completed successfully"
            } | Select-Object ComputerName,Manufacturer,Model,SerialNumber,Messages
        }
        Catch {
            $Msg = "WMI query failed"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += ";  $ErrorDetails"}
            
            New-Object PSObject -Property @{
                ComputerName = $Env:ComputerName
                Manufacturer = "Error"
                Model        = "Error"
                SerialNumber = "Error"
                Messages     = $Msg
            } | Select-Object ComputerName,Manufacturer,Model,SerialNumber,Messages
        }

    } #end scriptblock

    #endregion Scriptblock for Invoke-Command

    #region Functions

    Function Test-WinRM{
        Param($Computer)
        $Param_WSMAN = @{
            ComputerName   = $Computer
            Credential     = $Credential
            Authentication = "Kerberos"
            ErrorAction    = "Silentlycontinue"
            Verbose        = $False
        }
        Try {
            If (Test-WSMan @Param_WSMAN) {$True}
            Else {$False}
        }
        Catch {$False}
    }

    Function Test-Ping{
        Param($Computer)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
    }

    #endregion Functions

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Write-Progress
    $Activity = "Invoke scriptblock to return Windows hardware details"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as remote PSJob"
    }
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    # Parameters for Invoke-Command
    $ConfirmMsg = $Activity
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName   = ""
        Authentication = "Kerberos"
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.Add("AsJob",$True)
        $Param_IC.Add("JobName",$Null)
    }
    
    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}


} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        $Computer = $Computer.Trim()
        $Current ++ 

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = $Computer
        
        [switch]$Continue = $False

        Switch ($ConnectionTest) {
            Default {$Continue = $True}
            Ping {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Ping computer"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "Ping failure"
                            $Host.UI.WriteErrorLine("[$Computer] $Msg")
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        $Host.UI.WriteErrorLine("[$Computer] $Msg")
                    }
                }
                Else {$Continue = $True}
            }
            WinRM {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Test WinRM connection"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-WinRM -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "WinRM failure"
                            $Host.UI.WriteErrorLine("[$Computer] $Msg")
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        $Host.UI.WriteErrorLine("[$Computer] $Msg")
                    }
                }
                Else {$Continue = $True}
            }        
        }

        If ($Continue.IsPresent) {
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Msg = "Invoke command"
                    If ($AsJob.IsPresent) {$Msg += " as PSJob"}
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Jobs += $Job
                    }
                    Else {
                        $Output = Invoke-Command @Param_IC
                        $Output | Select -Property * -ExcludeProperty PSComputerName,RunspaceID
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine("[$Computer] $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                $Host.UI.WriteErrorLine("[$Computer] $Msg")
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {

        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            Write-Verbose $Msg
            $Jobs | Get-Job
            
        }
        Else {
            $Msg = "No jobs created"
            Write-Warning $Msg
        }
    } #end if AsJob

    $Msg = "END    : $Activity"
    $FGColor = "Green"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}
}

} # end Get-PKWindowsHardware
