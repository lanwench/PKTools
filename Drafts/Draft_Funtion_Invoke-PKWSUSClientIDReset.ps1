#Requires -version 3
Function Invoke-PKWSUSClientIDReset {
<# 
.SYNOPSIS
    Forces a WSUS client reset--including the client ID--interactively, or as a psjob

.DESCRIPTION
    Forces a WSUS client reset--including the client ID--interactively, or as a psjob
    Based on Boe Prox's original
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Funtion_Invoke-PKWSUSClientIDReset.ps1
    Created : 2018-12-05
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-12-05 - Created script

.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Invoke-WSUSClientFix-fd29e1a8

.PARAMETER ComputerName
    Name of computer to do cool thing on; separate multiple names with commas

.PARAMETER Credential
    Valid credentials on target computer (default is current user credentials)

.PARAMETER AsJob
   Invoke command as a job

.PARAMETER WaitForJob
    If running as a PSJob, wait for completion and return results

.PARAMETER JobWaitTimeout
    If WaitForJob, timeout in seconds for job output

.PARAMETER SkipConnectionTest
    Don't test WinRM connectivity before submitting command

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Invoke-PKWSUSClientIDReset -ComputerName server999 -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        ComputerName          {server999}                           
        Verbose               True                                     
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        WaitForJob            False                                    
        JobWaitTimeout        10                                       
        SkipConnectionTest    False                                    
        SuppressConsoleOutput False                                    
        ScriptName            Invoke-PKWSUSClientIDReset               
        ScriptVersion         1.0.0                                    

        Action: Force WSUS client ID reset
        VERBOSE: server999
        VERBOSE: Test WinRM connection

        ComputerName : SERVER999
        IsSuccess    : True
        Actions      : {Stoped wuauserv service, Connected to registry, Deleted registry value 'SusClientID', Deleted registry 
                       value 'SusClientIdValidation'...}
        Warnings     : {}
        Errors       : {}

.EXAMPLE
    PS C:\> $Arr | Invoke-PKWSUSClientIDReset -Credential $Credential -AsJob -WaitForJob -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Credential            System.Management.Automation.PSCredential
        AsJob                 True                                     
        WaitForJob            True                                     
        Verbose               True                                     
        ComputerName                                                   
        JobWaitTimeout        10                                       
        SkipConnectionTest    False                                    
        SuppressConsoleOutput False                                    
        ScriptName            Invoke-PKWSUSClientIDReset               
        ScriptVersion         1.0.0                                    

        Action: Force WSUS client ID reset as remote PSJob (wait 10 second(s) for job out
        put)

        VERBOSE: computer-7
        VERBOSE: Test WinRM connection
        VERBOSE: Job ID 441: WSUSReset_computer-7
        VERBOSE: foo
        VERBOSE: Test WinRM connection
        ERROR: WinRM connection failed on foo
        VERBOSE: sqlserver
        VERBOSE: Test WinRM connection
        Operation cancelled by user sqlserver
        VERBOSE: devwebserver9
        VERBOSE: Test WinRM connection
        VERBOSE: Job ID 443: WSUSReset_devwebserver9
        2 job(s) created (waiting 10 second(s) for output)


        ComputerName  : COMPUTER-7
        IsSuccess     : True
        RestartNeeded : True
        Actions       : {Stoped wuauserv service, Connected to registry, Deleted folder 'C:\Windows\SoftwareDistribution' 
                        and all subdirectories, Started wuauserv service...}
        Warnings      : {}
        Errors        : {}

        ComputerName  : DEVWEBSERVER9
        IsSuccess     : True
        RestartNeeded : True
        Actions       : {Stoped wuauserv service, Connected to registry, Deleted folder 'C:\Windows\SoftwareDistribution' 
                        and all subdirectories, Started wuauserv service...}
        Warnings      : {}
        Errors        : {}
        
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
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Hostname or FQDN of computer (separate multiple computers with commas)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Valid credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory = $False,
        HelpMessage = "Run as remote PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory=$False,
        HelpMessage="Wait for job completion"
    )]
    [Switch] $WaitForJob,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory=$False,
        HelpMessage="Timeout in seconds to wait for job results (default 10)"
    )]
    [int] $JobWaitTimeout = 10,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Don't test WinRM connectivity before invoking command"
    )]
    [Switch] $SkipConnectionTest,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)


Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
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
    
    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    $Results = @()
    
    # Scriptblock for invoke-command
    $SCriptBlock = {
        $Output = New-Object PSObject -Property ([ordered]@{
            ComputerName  = $env:COMPUTERNAME
            IsSuccess     = "Error"
            RestartNeeded = "Error"
            Actions       = "Error"
            Warnings      = "Error"
            Errors        = "Error"
        })

        $ErrMessages = @()
        $Actions = @()
        $Warnings = @()

        # Stop Windows service
        Try {
            $WUAUServ = Get-Service -Name wuauserv -EA Stop
            $Null = Stop-Service -InputObject $WUAUServ -EA Stop
            $Msg = "Stoped wuauserv service"
            $Actions += $Msg
            Write-Verbose $Msg
        }
        Catch {
            $Msg = "Service wuauserv failed to stop"
            $Msg += "`n$($_.Exception.Message)"
            $ErrMessages += $Msg
            $Host.UI.WriteErrorLine($Msg)
        }
                
        # Open registry
        Try {
            $reghive = [microsoft.win32.registryhive]::LocalMachine
            $Reg = [microsoft.win32.registrykey]::OpenBaseKey($reghive,"Default")

            # Connect to WSUS client reg keys
            $WSUSReg1 = $Reg.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\WindowsUpdate',$True)
            $WSUSReg2 = $Reg.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update',$True)

            $Msg = "Connected to registry"
            $Actions += $Msg
            Write-Verbose $Msg

            $RegVal1 = "SusClientID","SusClientIdValidation","PingID","AccountDomainSid"
            $RegVal2 = "LastWaitTimeout","DetectionStartTimeout","NextDetectionTime","Austate"
            Foreach ($Val in $RegVal1) {
                If (-Not [string]::IsNullOrEmpty($WSUSReg1.GetValue($Val))) {
                    Try {
                        $Null = $WSUSReg1.DeleteValue($Val)
                        $Msg = "Deleted registry value '$Val'"
                        $Actions += $Msg
                        Write-Verbose $Msg
                    }
                    Catch {
                        $Msg = "Error deleting registry value '$Val'"
                        $Msg += "`n$($_.Exception.Message)"
                        $ErrMessages += $Msg
                        $Host.UI.WriteErrorLine($Msg)
                    }
                }
            }
            $Val = $Null 
            Foreach ($Val in $RegVal2) {
                If (-Not [string]::IsNullOrEmpty($WSUSReg1.GetValue($Val))) {
                    Try {
                        $Null = $WSUSReg2.DeleteValue($Val)
                        $Msg = "Deleted registry value '$Val'"
                        $Actions += $Msg
                        Write-Verbose $Msg
                    }
                    Catch {
                        $Msg = "Error deleting registry value '$Val'"
                        $Msg += "`n$($_.Exception.Message)"
                        $ErrMessages += $Msg
                        $Host.UI.WriteErrorLine($Msg)
                    }
                }
            }
        }
        Catch {
            $Msg = "Registry connection error"
            $Msg += "`n$($_.Exception.Message)"
            $ErrMessages += $Msg
            $Host.UI.WriteErrorLine($Msg)
        }
            
        # Remove software distro folder & subfolders
        Try {
            $Null = Remove-Item "$Env:WinDir\SoftwareDistribution" -Recurse -Force -Confirm:$False -EA Stop                                                                                     
            $Msg = "Deleted folder '$Env:WinDir\SoftwareDistribution and all subdirectories'"
            $Actions += $Msg
            Write-Verbose $Msg

        } Catch {
            $Msg = "Error deleting folder '$Env:WinDir\SoftwareDistribution'"
            $Msg += "`n$($_.Exception.Message)"
            $ErrMessages += $Msg
            $Host.UI.WriteErrorLine($Msg)
        }
                
        # Start WUAUServ
        Try {
            $Null = Start-Service -InputObject $WUAUServ -EA SilentlyContinue
            $Msg = "Started wuauserv service"
            $Actions += $Msg
            Write-Verbose $Msg
        }
        Catch {
            $Msg = "Service wuauserv failed to start"
            $Msg += "`n$($_.Exception.Message)"
            $Warnings += $Msg
            Write-Warning $Msg
        }
                
        # Sent reauth/detectnow
        Try {
            $Null = Invoke-Expression -Command "wuauclt /resetauthorization /detectnow"
            $Msg = "Submitted reauthorization / detectnow to WSUS server"
            $Actions += $Msg
            Write-Verbose $Msg
        } Catch {
            $Msg = "Invocation of WSUS reauthorization request failed"
            $Msg += "`n$($_.Exception.Message)"
            $Warnings += $Msg
            Write-Warning $Msg
        }
            
        $Output.Warnings = $Warnings
        $Output.Actions = $Actions
        $Output.Errors = $ErrMessages

        If ($ErrMessages.Count -gt 0) {
            $Output.IsSuccess = $False
        }
        Else {
            $Output.IsSuccess = $True
            $Output.RestartNeeded = $True
        }

        Write-Output $Output
    } #end scriptblock

    #region Splats

    # Splat for Write-Progress
    $Activity = "Force WSUS client ID reset"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as remote PSJob"
        If ($WaitForJob.IsPresent) {$Activity = "$Activity (wait $JobWaitTimeout second(s) for job output)"}
    }
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    # Parameters for Test-WSMan
    $Param_WSMAN = @{}
    $Param_WSMAN = @{
        ComputerName   = ""
        Credential     = $Credential
        Authentication = "Kerberos"
        ErrorAction    = "Silentlycontinue"
        Verbose        = $False
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
        $Param_IC.AsJob = $True
        $Param_IC.JobName = $Null
        $JobPrefix = "WSUSReset"
    }
    
    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

    $Host.UI.WriteLine()
    $Msg = "Target computers will need to be restarted after successful run"

} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        $Computer = $Computer.Trim()
        $Msg = $Computer
        Write-Verbose $Msg

        $Current ++ 

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.CurrentOperation = $Computer
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = "$percentComplete%"
        Write-Progress @Param_WP
        
        [switch]$Continue = $False

        # If we're testing WinRM
        If (-not $SkipConnectionTest.IsPresent) {
                    
            $Msg = "Test WinRM connection"
            Write-Verbose $Msg

            If ($PSCmdlet.ShouldProcess($Computer,$Msg)) {

                $Param_WSMan.computerName = $Computer
                If ($Null = Test-WSMan @Param_WSMan ) {
                    $Continue = $True
                }
                Else {
                    $Msg = "WinRM connection failed on $Computer"
                    #If ($ErrorDetails = [regex]:: match($_.Exception.Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()) {$Msg = "$Msg`n$ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                }
            }
            Else {
                $Msg = "WinRM connection test cancelled by user"
                $Host.UI.WriteErrorLine("$Msg on $Computer")
            }
        } 
        Else {
            $Continue = $True
        }
        
        If ($Continue.IsPresent) {
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Msg = "Job ID $($Job.ID): $($Job.Name)"
                        Write-Verbose $Msg
                        $Jobs += $Job
                    }
                    Else {
                        $Results += Invoke-Command @Param_IC
                    }
                }
                Catch {
                    $Msg = "Operation failed on $Computer"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user on $Computer"
                $Host.UI.WriteErrorLine($Msg)
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {

        If ($Jobs.Count -gt 0) {
        
            If ($WaitForJob.IsPresent) {
                $Msg = "$($Jobs.Count) job(s) created (waiting $JobWaitTimeout second(s) for output)"
                $Activity = $Msg
                Write-Progress -Activity $Activity
                $FGColor = "Yellow"
                If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
                Else {Write-Verbose $Msg}

                Start-Sleep 5

                Try {
                    $Results = @()
                    If ($NotComplete = ($Jobs | Get-Job | Where-Object {(@("Failed","Running") -contains $_.State)})) {
                        $Msg = "Not all jobs completed within the $JobWaitTimeout-second timeout period. Please run 'Get-Job -Id #'`n"
                        Write-Warning $Msg
                        Write-Verbose ($Notcomplete | Get-Job @StdParams | Out-String)
                    }
                    If (($Results = $Jobs | Get-Job | Wait-Job -Timeout $JobWaitTimeout | Receive-Job -ErrorAction SilentlyContinue -Verbose:$False).Count -gt 0) {
                        Write-Output ($Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID)
                    }
                    Else {
                        $Msg = "No job output returned (try -IncludeNoMatch to return data for computers where no matching services were found)"
                        Write-Warning $Msg
                    }
                }
                Catch {
                    $Msg = "Please check job output manually using 'Get-Job -Id #' "
                    $ErrorDetails = $_.Exception.Message
                    $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
                }
            }
            Else {
                $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
                Write-Verbose $Msg
                $Jobs | Get-Job
            }
        }
        Else {
            $Msg = "No jobs created"
            $Host.UI.WriteErrorLine($Msg)
        }
    } #end if AsJob

    Else {
        If ($Results.Count -eq 0) {
            $Msg = "No results found"
            $Host.UI.WriteErrorLine($Msg)
        }
        Else {
            Write-Output ($Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID)
        }
    }

}

} # end Invoke-PKWSUSClientIDReset
