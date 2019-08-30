#Requires -version 3
Function Get-PKTaniumConfig {
<# 
.SYNOPSIS
    Invokes a scriptblock to get Tanium client config details, interactively or as a PSJob

.DESCRIPTION
    Invokes a scriptblock to get Tanium client config details, interactively or as a PSJob
    Accepts pipeline input
    Optionally tests connectivity to remote computers before invoking scriptblock
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKTaniumConfig.ps1
    Created : 2019-08-20
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-08-20 - Created script

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER Credential
    Valid credentials on target (default is passthrough)

.PARAMETER Authentication
    Authentication protocol: Basic, CredSSP, Negotiate, Kerberos, or None (default is Negotiate)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'Tanium')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command: WinRM, ping, or none (default is WinRM)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> 

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
        HelpMessage = "Authentication protocol: Basic, CredSSP, Negotiate, Kerberos, or None (default is Negotiate)"
    )]
    [ValidateSet("None","CredSSP","Basic","Negotiate","Kerberos")]
    [string] $Authentication = "Negotiate",

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
        HelpMessage = "Prefix for job name (default is 'Tanium')"
    )]
    [String] $JobPrefix = "Tanium",

    [Parameter(
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM with Kerberos, ping, or none (default is WinRM)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

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
    $ProgressPreference    = "Continue"
    
    #region Scriptblock for Invoke-Command
    $ScriptBlock = {
    
        $Messages = @()
        $Output = New-Object PSObject -Property @{
            ComputerName    = $Env:ComputerName
            FQDN            = "Error"
            TaniumInstalled = "Error"
            Path            = "Error"
            Version         = "Error"
            ServerName      = "Error"
            ServerNameList  = "Error"
            ServerPort      = "Error"
            FirstInstall    = "Error"
            ComputerID      = "Error"
            Tags            = "Error"
            ServiceState    = "Error"
            Messages        = "Error"
        }
        $Select = 'ComputerName','FQDN','TaniumInstalled','Path','Version','ServerName','ServerNameList','ServerPort','FirstInstall','ComputerID','Tags','ServiceState','Messages'
    
        If ($FQDN = (Get-WmiObject Win32_ComputerSystem | Select  @{N="FQDN";E={"$($_.Name).$($_.Domain)"}}).FQDN) {
            $Output.FQDN = $FQDN
        }
        Else {
            $Output.FQDN = "-"
        }

        # Get the client install
        $Reg_TaniumClient = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\TaniumClient.exe"
        If ($TaniumClient = Get-Item -Path $Reg_TaniumClient -EA SilentlyContinue) {
            $TaniumPath = ($TaniumClient | Get-ItemProperty)."(default)"
            $Output.TaniumInstalled = $True
            $Output.Path = $TaniumPath

            # Get the client config
            $Reg_TaniumConfig = "HKLM:\SOFTWARE\WOW6432Node\Tanium\Tanium Client"
            If ($TaniumConfig = Get-Item -Path $Reg_TaniumConfig -EA SilentlyContinue | Get-ItemProperty) {
                
                #Get-Date $TaniumConfig.FirstInstall 
                #$Install = [datetime]::parseexact($TaniumConfig.FirstInstall, 'dd/MM/yyyy HH:mm:ss', $null).ToString('yyyy-MM-dd HH:mm:ss')
                
                $Output.Version        = $TaniumConfig.Version
                $Output.ServerName     = $TaniumConfig.ServerName
                $Output.ServerNameList = $TaniumConfig.ServerNameList
                $Output.ServerPort     = $TaniumConfig.ServerPort
                $Output.FirstInstall   = $TaniumConfig.FirstInstall
                $Output.ComputerID     = $TaniumConfig.ComputerID
            }

            # Get the tags
            $Reg_TaniumTag = "HKLM:\SOFTWARE\Wow6432Node\Tanium\Tanium Client\Sensor Data\"
            If ($TaniumTags = Get-ChildItem -Path $Reg_TaniumTag -EA SilentlyContinue  | Where-Object {$_.Name -match "Tags"} | Select -ExpandProperty Property) {
                $Output.Tags = $TaniumTags
            }
            Else {
                $Output.Tags = "-"
                $Messages += "No tags found"
            }
            # Get the Windows service
            If ($TaniumService = Get-Service -Name "Tanium Client" -EA SilentlyContinue) {
                $Output.ServiceState = $TaniumService.Status
            }
            Else {
                $Messages += "Tanium Client service not found"
            }
    
        } # end if client found
        Else {
            $Output.TaniumInstalled = $False
            $Messages += "Tanium Client not found in registry"
        }

        $Output.Messages = $Messages -join("`n")
        Write-Output $Output | Select-Object $Select

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
        Param([Parameter(ValueFromPipeline)]$Message)#,[switch]$Quiet = $Quiet)
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
    $Activity = "Invoke scriptblock to get Tanium Client configuration"
    If ($AsJob.IsPresent) {
        $Activity += " (as job)"
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
        Authentication = $Authentication
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

    #region Failure log

    [string[]]$Failures = @()

    #endregion Failure log

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

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "Ping failure"
                            "[$Computer] $Msg" | Write-MessageError
                            $Failures += $Computer
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                        $Failures += $Computer
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

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-WinRM -Computer $Computer) {
                            $Continue = $True
                        }
                        Else {
                            $Msg = "WinRM failure"
                            "[$Computer] $Msg" | Write-MessageError
                            $Failures += $Computer
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                        $Failures += $Computer
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
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Jobs += $Job
                    }
                    Else {
                        Invoke-Command @Param_IC
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[$Computer] $Msg" | Write-MessageError
                    $Failures += $Computer
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor Cyan
                $Failures += $Computer
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

    If ($Failures.Count -gt 0) {
        $Msg = "Operation failed or cancelled on $($Failures.Count) computers:`n$($Failures -join("`n"))"
        Write-Warning $Msg
    }

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

