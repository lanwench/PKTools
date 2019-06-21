#Requires -version 3
Function Get-PKTaniumActionLog {
<# 
.SYNOPSIS
    Invokes a scriptblock to parse content from the Tanium client action log, returning a PSObject

.DESCRIPTION
    Invokes a scriptblock to parse content from the Tanium client action log, returning a PSObject
    Accepts pipeline input
    Optionally tests connectivity to remote computers before invoking scriptblock
    Returns a PSobject or PSJob

.NOTES        
    Name    : Get-PKTaniumActionLog.ps1
    Created : 2019-06-19
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-06-19 - Created script

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER CustomFilePath
    Absolute path to Tanium client action logfile (default is to look for action-history text file in c:\Program Files (x86)\Tanium\Tanium Client\Logs)

.PARAMETER Credential
    Valid credentials on target (default is passthrough)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'TaniumActivity')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM with Kerberos, ping, or none (default is WinRM)

.PARAMETER NoProgress
    Don't display progress bar (can improve performance with large files)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> Get-PKTaniumActionLog -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        Verbose        True                                     
        ComputerName   WORKSTATION11                          
        CustomFilePath                                          
        Credential     System.Management.Automation.PSCredential
        AsJob          False                                    
        JobPrefix      TaniumActivity                           
        ConnectionTest WinRM                                    
        NoProgress     False                                    
        Quiet          False                                    
        ScriptName     Get-PKTaniumActionLog                  
        ScriptVersion  1.0.0                                    
        PipelineInput  False                                    

        BEGIN  : Get Tanium Client action log content

        [WORKSTATION11] Invoke command

        Computername : WORKSTATION11
        Date         : 2019-06-21 11:28:13 AM
        ActionID     : 1070371
        Action       : Deploy - Maintenance Window 1
        Message      : completed in 0.161 seconds with exit code 0

        Computername : WORKSTATION11
        Date         : 2019-06-21 11:28:13 AM
        ActionID     : 1070371
        Action       : Deploy - Maintenance Window 1
        Message      : executing: cmd /c xcopy /y maintenance_window_1.json ..\..\Tools\Deploy\maintenance_windows\

        Computername : WORKSTATION11
        Date         : 2019-06-21 11:27:27 AM
        ActionID     : 1070369
        Action       : Patch - Patch Lists - Windows
        Message      : completed in 0.237 seconds with exit code 0

        Computername : WORKSTATION11
        Date         : 2019-06-21 11:27:27 AM
        ActionID     : 1070369
        Action       : Patch - Patch Lists - Windows
        Message      : executing: cmd /c xcopy /y patch-lists ..\..\Patch\

        Computername : WORKSTATION11
        Date         : 2019-06-21 11:25:51 AM
        ActionID     : 1070365
        Action       : Deploy - Profile 2
        Message      : completed in 1.123 seconds with exit code 0

        <snip>

        END    : Get Tanium Client action log content

.EXAMPLE
    PS C:\> Get-PKTaniumActionLog -ComputerName sqlserver.domain.local -Credential $Credential -NoProgress -Quiet -AsJob

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        17     TaniumActivi... RemoteJob       Completed     True            sqlserver.domain.... ...                      


        C:\> $ Get-Job 17 | Receive-Job


        Computername   : SQLSERVER
        Date           : -
        ActionID       : -
        Action         : -
        Message        : No Tanium Client action log file found in C:\Program Files*\Tanium\Tanium Client
        PSComputerName : sqlserver.domain.local
        RunspaceId     : 8f66ec17-900d-4719-910c-6344abccb5ea


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
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more computer names (default is local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [object[]]$ComputerName,

    [Parameter(
        HelpMessage = "Absolute path to Tanium client action logfile (default is to look for action-history* text file in c:\Program Files (x86)\Tanium\Tanium Client\Logs)"
    )]
    [ValidateNotNullOrEmpty()]
    [string] $CustomFilePath,

    [Parameter(
        HelpMessage = "Valid credentials on target (default is passthrough)"
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
        HelpMessage = "Prefix for job name (default is 'TaniumActivity')"
    )]
    [String] $JobPrefix = "TaniumActivity",

    [Parameter(
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM with Kerberos, ping, or none (default is WinRM)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

    [Parameter(
        HelpMessage = "Don't display progress bar (can improve performance with large files)"
    )]
    [Switch] $NoProgress,

    [Parameter(
        HelpMessage = "Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName
    }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    Switch ($NoProgress) {
        $True  {$ProgressPreference = "SilentlyContinue"}
        $False {$ProgressPreference = "Continue"}
    }
    
    
    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        
        Param($CustomFilePath)

        If ($CustomFilePath) {
            If (Test-Path $CustomFilePath -PathType Container -EA SilentlyContinue) {
                Try {
                    $InputFile = Get-ChildItem $CustomFilePath -Include *.txt -EA SilentlyContinue | Where-Object {$_.Name -match "action-history.\.txt"}
                }
                Catch {
                    $Msg = "No Tanium client activity logfile found in path '$CustomFilePath"
                }
            }
            Else {
                $Msg = "Invalid path '$CustomFilePath"
            }
        }
        Else {
            Try {
                $TaniumCommandPath = "$($Env:SystemDrive)\Program Files*\Tanium\Tanium Client\taniumclient.exe"
                If ($Parent = Get-Command -Name $TaniumCommandPath | Select-Object $_.Path | Split-Path -Parent) {
                    If ($ActionFile = $Parent | Get-ChildItem -Recurse -Include *.txt -EA SilentlyContinue | Where-Object {$_.Name -match "action-history.\.txt"}) {
                        $InputFile = $ActionFile.FullName
                    }
                    Else {
                        $Msg = "No Tanium Client action log file found in $($TaniumCommandPath | Select-Object $_.Path | Split-Path -Parent)"
                    }
                }
                Else {
                    $Msg = "Tanium Client path not found"
                }
            }
            Catch {
                $Msg = "Something horrible has happened!"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            }
        }
        Try {
            If ($Null = Get-Item $InputFile -EA SilentlyContinue) {

                Try { 

                    If ($LogContent = Get-Content $InputFile -Raw -ErrorAction SilentlyContinue) {       
                    
                        # Thanks for the regex help, Jeffery Hicks!
                        [regex]$rx = "(?<Date>\d{4}.*Z).*Action\s(?<ActionID>\d+).*\s+\[(?<Action>.*)\]:\s+(?<Message>.*)"
                    
                        If ($M = $rx.Matches($LogContent)) {
                            $Output = @()
                            $Output += Foreach ($item in $m) {
                                [pscustomobject]@{
                                    Computername = $env:Computername
                                    Date         = $item.groups[1].Value -as [datetime]
                                    ActionID     = $item.groups[2].value
                                    Action       = $item.groups[3].value
                                    Message      = $item.groups[4].value
                                }
                            }
                        }
                        Else {
                            $Msg = "No match found in $InputFile"
                        }
                    }
                    Else {
                        $Msg = "Failed to get content in $InputFile"
                    }
                }
                Catch {
                    $Msg = "Something horrible has happened!"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                }
            }
        }
        Catch {}
            
        If (-not $Output) {
            $Output = [pscustomobject]@{
                Computername = $env:Computername
                Date         = "-"
                ActionID     = "-"
                Action       = "-"
                Message      = $Msg
            }
        }
        Write-Output $Output | Select-Object ComputerName,Date,ActionID,Action,Message
    } #end scriptblock

    #endregion Scriptblock for Invoke-Command

    #region Functions

    # Function to write a console message or a verbose message
    Function Write-MessageInfo {
        Param([Parameter(ValueFromPipeline)]$Message,$FGColor,[switch]$Title)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {
            If ($Title.IsPresent) {$Message = "`n$Message`n"}
            $Host.UI.WriteLine($FGColor,$BGColor,"$Message")
        }
        Else {Write-Verbose "$Message"}
    }

    # Function to write an error or a verbose message
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("$Message")}
        Else {Write-Verbose "$Message"}
    }

    # Function to test WinRM connectivity
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

    # Function to test ping connectivity
    Function Test-Ping{
        Param($Computer)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
    }

    #endregion Functions

    #region Splats

    # Splat for Write-Progress
    $Activity = "Get Tanium Client action log content"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity (as job)"
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
        ComputerName   = $Null
        Authentication = "Kerberos"
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($CurrentParams.CustomFilePath) {
        $Param_IC.Add("ArgumentList",$CustomFilePath)
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.Add("AsJob",$True)
        $Param_IC.Add("JobName",$Null)
    }
    
    #endregion Splats

    # Console output
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title


} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
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
        
        $Current ++ 
        $Param_WP.PercentComplete = ($Current/$Total* 100)
        $Param_WP.Status = $Computer
        
        [switch]$Continue = $False

        Switch ($ConnectionTest) {
            Default {$Continue = $True}
            Ping {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Ping computer"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n`n`t$Msg`n")) {
                        If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "Ping failure"
                            "[$Computer] $Msg" | Write-MessageError
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                    }
                }
                Else {$Continue = $True}
            }
            WinRM {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Test WinRM connection"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n`n`t$Msg`n`n")) {
                        If ($Null = Test-WinRM -Computer $Computer) {
                            $Continue = $True
                        }
                        Else {
                            $Msg = "WinRM failure"
                            "[$Computer] $Msg" | Write-MessageError
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                    }
                }
                Else {$Continue = $True}
            }        
        }

        If ($Continue.IsPresent) {
            
            $Msg = "Invoke command"
            If ($AsJob.IsPresent) {$Msg += " as PSJob"}
            "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP

            If ($PSCmdlet.ShouldProcess($Computer,"`n`n`t$Activity`n`n")) {
            
                Try {
                    
                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Jobs += $Job
                    }
                    Else {
                        Invoke-Command @Param_IC | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceId
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[$Computer] $Msg" | Write-MessageError
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

    If ($AsJob.IsPresent -and ($Jobs.Count -gt 0)) {

        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output`n"
            "$Msg" | Write-MessageInfo -FGColor White -Title
            $Jobs | Get-Job
            
        }
        Else {
            $Msg = "No jobs created"
            $Msg | Write-MessageError
        }
    } #end if AsJob


    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title


}

} # end function

