#Requires -version 3
Function Get-PKWindowsProductKey {
<# 
.SYNOPSIS
    Uses WMI to retrieve a Windows product key on a local or remote computer, interactively or as a PSJob

.DESCRIPTION
    Uses WMI to retrieve a Windows product key on a local or remote computer, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsProductKey.ps1
    Created : 2019-02-25
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-02-25 - Created script

.LINK
    https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/managing-windows-license-key

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
    PS C:\Users\pkingsley\Git\Modules\PKTools> Get-PKWindowsProductKey -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Verbose               True                                     
        ComputerName          {WORKSTATION14}                        
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        JobPrefix             Job                                      
        ConnectionTest        None                                     
        SuppressConsoleOutput False                                    
        ScriptName            Get-PKWindowsProductKey                  
        ScriptVersion         1.0.0                                    

        ACTION : Invoke scriptblock to return Windows product key
        VERBOSE: [WORKSTATION14] Invoke command

        OperatingSystem                 Version        ProductKey                   
        ---------------                 -------        ----------                   
        Microsoft Windows 10 Enterprise 10.0.17134.471 V****-9****-6****-B****-P****

.EXAMPLE
    PS C:\> $Arr | Get-PKWindowsProductKey -Credential $Credential -AsJob -ConnectionTest WinRM -SuppressConsoleOutput -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Credential            System.Management.Automation.PSCredential
        AsJob                 True                                     
        ConnectionTest        WinRM                                    
        SuppressConsoleOutput True                                     
        ComputerName          {WORKSTATION14}                        
        JobPrefix             Job                                      
        ScriptName            Get-PKWindowsProductKey                  
        ScriptVersion         1.0.0                                    

        VERBOSE: ACTION : Invoke scriptblock to return Windows product key as remote PSJob
        VERBOSE: [ops-pkbastion-1.domain.local] Test WinRM connection
        VERBOSE: [ops-pkbastion-1.domain.local] WinRM connection test successful
        VERBOSE: [ops-pkbastion-1.domain.local] Invoke command as PSJob
        VERBOSE: [ops-wsus-201.domain.local] Test WinRM connection
        VERBOSE: [ops-wsus-201.domain.local] WinRM connection test successful
        VERBOSE: [ops-wsus-201.domain.local] Invoke command as PSJob
        VERBOSE: 2 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output
        
        <snip>

        PS C:\> Get-Job | Receive-Job

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        18     Job_ops-pkba... RemoteJob       Running       True            ops-pkbastion-1.d... ...                      
        20     Job_ops-wsus... RemoteJob       Running       True            ops-wsus-201.doma... ...                      



        
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
        #Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName = $env:COMPUTERNAME,

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
        HelpMessage = "Prefix for job name (default is 'WinKey')"
    )]
    [String] $JobPrefix = "WinKey",

    [Parameter(
        HelpMessage="Test to run prior to Invoke-Command: WinRM, Ping, or None (default None)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "None",

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
        Get-WmiObject -Class SoftwareLicensingService | Select PSComputerName,
        @{N="OperatingSystem";E={(Get-WMIObject -Class Win32_OperatingSystem).Caption}},
        Version,
        @{N="ProductKey";E={$_.OA3xOriginalProductKey}}
    }

    #endregion Scriptblock for Invoke-Command

    #region Functions

    Function Test-WinRM{
        Param($Computer,$Credential)
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
    $Activity = "Invoke scriptblock to return Windows product key"
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
    $Msg = "START   : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
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
                    $ConfirmMsg = "`n`n`t$Msg`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                    #If ($ConnectionTest) {
                        If ($Null = Test-WinRM -Computer $Computer -Credential $Credential) {
                            $Continue = $True
                            $Msg = "WinRM connection test successful"
                            Write-Verbose "[$Computer] $Msg"
                        }
                        Else {
                            $Msg = "WinRM connection test failure"
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
                        $Results = Invoke-Command @Param_IC 
                        Write-Output $Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR : [$Computer] $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                Write-Warning "[$Computer] $Msg"
                $Host.UI.WriteErrorLine("ERROR  : [$Computer] $Msg")
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {

        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"END    : $Msg")}
            Else {Write-Verbose "END    : $Msg"}
            $Jobs | Get-Job
        }
        Else {
            $Msg = "No jobs created"
            Write-Warning "$Msg"
        }
    } #end if AsJob
}

} # end Get-PKWindowsProduct Key

