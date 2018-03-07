#requires -Version 3
Function Install-PKWindowsPythonPackage {
<#
.SYNOPSIS
    Installs a Python pip package via Invoke-Command, synchronously or in a PSJob

.DESCRIPTION
    Installs a Python pip package via Invoke-Command, synchronously or in a PSJOb
    Requires that Python/pip be installed on target computer
    If pip is not found, but Chocolatey is found, suggests installation via choco.exe
    If package is already found, exits unless -Force or -Upgrade is specified
    Accepts pipeline input
    Returns a PSObject or PSJob

.NOTES
    Name    : Function_Install-PKWindowsPIPPackage.ps1
    Created : 2017-12-01
    Version : 01.00.0000
    Author  : Paula Kingsley
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-12-01 - Created script

.PARAMETER ComputerName
    One or more target computers

.PARAMETER Package
    One or more Python pip package names (must be exact/valid)

.PARAMETER Upgrade
    Upgrade package if found

.PARAMETER Force
    Force package installation if found

.PARAMETER AsJob
    Run as remote PSJob

.PARAMETER Credential
    Valid admin credential on target computer

.PARAMETER SkipConnectionTest
    Don't test WinRM connectivity before running Invoke-Command

.PARAMETER SuppressConsoleOutput
    Suppress non-verbose/non-error console output

.EXAMPLE
    PS C:\> Install-PKWindowsPIPPackage -ComputerName localhost -Package fuzzywuzzy,elasticsearch -Force -Verbose

        VERBOSE: PSBoundParameters:

        Key                   Value
        ---                   -----
        ComputerName          {localhost}
        Package               {fuzzywuzzy, elasticsearch}
        Force                 True
        Verbose               True
        Upgrade               False
        AsJob                 False
        Credential            System.Management.Automation.PSCredential
        SkipConnectionTest    False
        SuppressConsoleOutput False
        PipelineInput         False
        ScriptName            Install-PKWindowsPIPPackage
        ScriptVersion         1.0.0

        Action: Test connection and install pip packages 'fuzzywuzzy', 'elasticsearch' (force)
        VERBOSE: localhost


        ComputerName : WORKSTATION
        Package      : fuzzywuzzy
        IsInstalled  : True
        Force        : True
        Upgrade      : False
        pipVersion   : pip 9.0.1 from c:\python36\lib\site-packages (python 3.6)
        Messages     : Successfully installed fuzzywuzzy-0.15.1

        ComputerName : WORKSTATION
        Package      : elasticsearch
        IsInstalled  : True
        Force        : True
        Upgrade      : False
        pipVersion   : pip 9.0.1 from c:\python36\lib\site-packages (python 3.6)
        Messages     : Successfully installed elasticsearch-6.0.0 urllib3-1.22

.EXAMPLE
    PS C:\> $Computers | Install-PKWindowsPIPPackage -Package fuzzywuzzy -Verbose

        VERBOSE: PSBoundParameters:

        Key                   Value
        ---                   -----
        Package               fuzzywuzzy
        Verbose               True
        ComputerName
        Upgrade               False
        Force                 False
        AsJob                 False
        Credential            System.Management.Automation.PSCredential
        SkipConnectionTest    False
        SuppressConsoleOutput False
        PipelineInput         True
        ScriptName            Install-PKWindowsPIPPackage
        ScriptVersion         1.0.0

        Action: Test connection and install pip package 'fuzzywuzzy'
        VERBOSE: sqlserver.domain.local
        VERBOSE: webserver.domain.local
        VERBOSE: foo.domain.local
        ERROR: Connection failure on foo.domain.local [The WinRM client cannot process the request.
        Default credentials with Negotiate over HTTP can be used only if the target machine is part of
        the TrustedHosts list or the Allow implicit credentials for Negotiate option is specified.]
        VERBOSE: exchangeserver.domain.local
        Operation cancelled by user on exchangeserver.domain.local

        ComputerName   : SQLSERVER
        Package        : fuzzywuzzy
        IsInstalled    : False
        Force          : False
        Upgrade        : False
        pipVersion     : pip 9.0.1 from C:\Python27\lib\site-packages (python 2.7)
        Messages       : Package 'fuzzywuzzy (0.15.1)' already installed; please specify -Force or -Upgrade

        ComputerName   : WEBSERVER
        Package        : fuzzywuzzy
        IsInstalled    : True
        Force          : False
        Upgrade        : False
        pipVersion     : pip 9.0.1 from c:\python36\lib\site-packages (python 3.6)
        Messages       : Successfully installed fuzzywuzzy-0.15.1

.EXAMPLE
    PS C:\> $ Install-PKWindowsPIPPackage -ComputerName devbox -Package treetime -Force -Credential $Credential -SkipConnectionTest -AsJob -Verbose

        VERBOSE: PSBoundParameters:

        Key                   Value
        ---                   -----
        ComputerName          {devbox}
        Package               treetime
        Force                 True
        Credential            System.Management.Automation.PSCredential
        SkipConnectionTest    True
        AsJob                 True
        Verbose               True
        Upgrade               False
        SuppressConsoleOutput False
        PipelineInput         False
        ScriptName            Install-PKWindowsPIPPackage
        ScriptVersion         1.0.0

        Action: Install pip package 'treetime' (force) as PSJob
        VERBOSE: devbox
        VERBOSE: Created job 33
        VERBOSE: 1 job(s) submitted; run 'Get-Job x | Wait-Job | Receive-Job'

        Id     Name        PSJobTypeName   State         HasMoreData     Location       Command
        --     ----        -------------   -----         -----------     --------       -------
        33     pip_devb... RemoteJob       Running       True            devbox.int...  ...

        [...]

        PS C:\> Get-Job 33 | Receive-job

        ComputerName   : DEVBOX
        Package        : treetime
        IsInstalled    : Error
        Force          : True
        Upgrade        : False
        pipVersion     : Error
        Messages       : pip.exe not found
                         Chocolatey detected; you can install pip via 'choco install pip -y''

.EXAMPLE
    PS C:\> Install-PKWindowsPIPPackage -ComputerName testvm -Package boguspackage -SuppressConsoleOutput

    ComputerName   : TESTVM
    Package        : boguspackage
    IsInstalled    : Error
    Force          : False
    Upgrade        : False
    pipVersion     : pip 9.0.1 from c:\python36\lib\site-packages (python 3.6)
    Messages       : No matching distribution found


#>
[Cmdletbinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(

    [Parameter(
        Mandatory = $True,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Name of computer (separate multiple computers with commas)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name","HostName","FQDN","DNSDomainName")]
    [String[]] $ComputerName,

    [Parameter(
        Mandatory = $True,
        Position = 1,
        HelpMessage = "One or more pip package names (must be exact!)"
    )]
    [ValidateNotNullOrEmpty()]
    [String[]] $Package,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Upgrade installation"
    )]
    [Switch] $Upgrade,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Force installation"
    )]
    [Switch] $Force,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Run as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Valid credentials on target computer(s)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty ,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Don't test WinRM connection before submitting command"
    )]
    [Switch] $SkipConnectionTest,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Suppress non-verbose console output"
    )]
    [Switch] $SuppressConsoleOutput
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    #$Source = $PSCmdlet.ParameterSetName
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

    # Preference
    $ErrorActionPreference = "Stop"

    # We can't do both, but need both parameters available
    If ($Force.IsPresent -and $Upgrade.IsPresent) {
        $Msg = "Please select either -Force or -Upgrade"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }

    # Output object if not running as job
    $Results = @()

    # Create scriptblock
    $Scriptblock = {

        Param($Package,$Force,$Upgrade)
        $Package = $Using:Package
        [switch]$Force = $Using:Force
        [switch]$Upgrade = $Using:Upgrade

        $StdParams = @{}
        $StdParams = @{
            ErrorAction = "Stop"
            Verbose     = $False
        }

        $Results = @()
        $InitialValue = "Error"
        $OutputTemplate = New-Object PSObject -Property @{
            ComputerName  = $Env:ComputerName
            Package       = $Package
            IsInstalled   = $InitialValue
            Force         = $Force
            Upgrade       = $Upgrade
            pipVersion    = $InitialValue
            PythonPath    = $InitialValue
            PythonVersion = $InitialValue
            Messages      = $InitialValue
        }
        $Select = "ComputerName","Package","IsInstalled","Force","Upgrade","pipVersion","PythonVersion","Messages"

        [switch]$Continue = $False

        # Check for Python
        Try {
            If (-not ($Python = Get-Command python -All -EA SilentlyContinue)) {
                $Msg = "Python not found"
                If ($Choco = Get-Command choco -EA SilentlyContinue) {
                    $Msg = "$Msg`nChocolatey detected; you can install Python via 'choco install python -y'"
                }
                $OutputTemplate.IsInstalled = $False
                $OutputTemplate.Messages = $Msg
                $Results += $OutputTemplate
            }
            Else {
                $OutputTemplate.PythonVersion = $Python.Version
                $OutputTemplate.PythonPath = $Python.Source

                If ($Python.Count -gt 1) {
                    $Msg = "Multiple versions of Python found; please adjust system path & `$env:PythonHome to the value you wish"
                    $OutputTemplate.IsInstalled = $False
                    $OutputTemplate.Messages = $Msg
                    $Results += $OutputTemplate
                }
                Else {
                    $Continue = $True
                }
            }
        }
        Catch {
            $Msg = "Python not found"
            If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
            $OutputTemplate.Messages = $Msg
            $Results += $OutputTemplate
        }

        # Check for pip
        If ($Continue.IsPresent) {

            $Continue = $False
            Try {
                If (-not ($PIP = Get-Command pip -EA SilentlyContinue)) {
                    $Msg = "pip.exe not found"
                    If ($Choco = Get-Command choco -EA SilentlyContinue) {
                        $Msg = "$Msg`nChocolatey detected; you can install pip via 'choco install pip -y'"
                    }
                    $OutputTemplate.Messages = $Msg
                    $Results += $OutputTemplate
                }
                Else {
                    $Continue = $True
                    $PipVer = Invoke-Expression -Command "pip --version"
                    $OutputTemplate.pipVersion = $($PipVer | Where-Object {$_} | Out-String).Trim()
                }
            }
            Catch {
                $Msg = "pip.exe not found"
                If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                $OutputTemplate.Messages = $Msg
                $Results = $OutputTemplate
            }
        }

        # If pip found, look for the packages and current installation state for each
        If ($Continue.IsPresent) {

            Foreach ($Pkg in $Package) {

                $continue = $False

                $Pkg = $Pkg.Trim()

                $Output = $OutputTemplate.PSObject.Copy()
                $Output.Package = $Pkg

                $Expression = "pip install $Pkg"

                Try {
                    $Null = Invoke-Expression -Command 'pip list 2>&1' -ErrorAction SilentlyContinue -OutVariable Test
                    If ($Exists = ($Test -match $Pkg)) {
                        If ($Force.IsPresent) {
                            $Expression = "$Expression --ignore-installed"
                            $Continue = $True
                        }
                        Elseif ($Upgrade.IsPresent) {
                            $Expression = "$Expression  --force-reinstall"
                            $Continue = $True
                        }
                        Else {
                            $Msg = "Package '$Exists' already installed; please specify -Force or -Upgrade"
                            $Output.IsInstalled = $False
                            $Output.Messages = $Msg
                            $Results += $Output
                        }
                    }
                    Else {
                        $Msg = "Verified package not currently installed"
                        $Output.Messages = $Msg
                        $Continue = $True
                    }
                }
                Catch {
                    $Msg = "Package installation status lookup failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                    $Output.IsInstalled = $False
                    $Output.Messages = $Msg
                    $Results += $Output
                }

                # If package isn't installed, or we are forcing install anyway
                If ($Continue.IsPresent) {

                    $Continue = $False

                    Try {

                        $Expression = "$Expression -v 2>&1"
                        $Null = Invoke-Expression -Command $Expression -OutVariable Install -ErrorAction SilentlyContinue

                        Switch -Regex ($Install) {
                            "error" {
                                $Output.Messages = $Install
                            }
                            "Requirement already satisfied" {
                                $Output.IsInstalled = $False
                                $Output.Messages = $($Install -match "Requirement already satisfied")
                            }
                            "Successfully installed" {
                                $Output.IsInstalled = $True
                                $Output.Messages = $($Install -match "successfully installed")
                            }
                            "No matching distribution found" {
                                $Output.Messages = $("No matching distribution found")
                            }
                        }
                    }
                    Catch {
                        $Msg = "Package installation failed"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                        $Output.IsInstalled = $False
                        $Output.Messages = $Msg
                    }

                    $Results += $Output

                } #end if continue

            } #end for each package

        } #end if pip found

        Write-Output $Results  | Select $Select

    } # end scriptblock

    #region Splats

    # Activity for write-progress & confirmation message
    If ($SkipConnectionTest.IsPresent) {[string]$Activity = "Test connection and install pip "}
    Else {[string]$Activity = "Install pip "}
    If ($Package.Count -gt 1) {
        $PkgStr = "'$($Package -join("', '"))'"
        $Activity += "packages $PkgStr"
    }
    Else {
        $Activity += "package '$Package'"
    }
    If ($Force.IsPresent) {$Activity += " (force)"}
    Elseif ($Upgrade.IsPresent) {$Activity += " (upgrade)"}

    $ConfirmMsg = "`n`n`t$Activity`n`n"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for write-progress
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        PercentComplete  = $Null
        CurrentOperation = $Null
        Status           = "Working"
    }

    # Splat for test-wsman
    $Param_WSMan = @{}
    $Param_WSMan = @{
        ComputerName   = $Null
        Authentication = "Kerberos"
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }

    # Splat for invoke-command
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName = ""
        Credential   = $Credential
        Scriptblock  = $Scriptblock
        ArgumentList = $Package,$Upgrade,$Force
        ErrorAction  = "Stop"
        Verbose      = $False
    }
    If ($AsJob.IsPresent) {
       $Activity += " as PSJob"
       $Param_IC.Add("AsJob",$True)
       $Param_IC.Add("JobName",$Null)
       $Jobs = @()
    }

    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}

Process {

    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {

        $Current ++

        Write-Verbose $Computer
        $Param_WP.CurrentOperation = $Computer
        $Param_WP.PercentComplete = ($Current / $Total * 100)
        Write-Progress @Param_WP

        [Switch] $Continue = $True
        If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {

            If (-not $SkipConnectionTest.IsPresent) {

                If ($Computer -eq $env:COMPUTERNAME) {
                    $Msg = "(Skipping connection test to local computer)"
                    Write-Verbose $Msg
                    $Continue = $True
                }
                Else {
                    Try {
                        $Param_WSMAN.ComputerName = $Computer
                        If ($Null = Test-WSMan @Param_WSMan ) {$Continue = $True}
                    }
                    Catch {
                        $Msg = "Connection failure on $Computer"
                        $ErrorDetails= [regex]:: match($_.Exception.Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()
                        $Host.UI.WriteErrorLine("ERROR: $Msg [$ErrorDetails]")
                        $Continue = $False
                    }
                }
            } #end if testing WinRM
            Else {$Continue = $True}

            If ($Continue.IsPresent) {

                $Param_IC.ComputerName = $Computer

                If ($AsJob.IsPresent) {
                    $Param_IC.JobName = "pip_$Computer"
                    $Job =  Invoke-Command @Param_IC
                    $Jobs += $Job
                    $Msg = "Created job $($Job.ID)"
                    Write-Verbose $Msg
                }
                Else {
                    $Results += Invoke-Command @Param_IC
                }
            }
        } #end if confirm
        Else {
            $Msg = "Operation cancelled by user on $Computer"
            $Host.UI.WriteErrorLine($Msg)
        } #end if cancelled
    } #end foreach computer

}
End {

    Write-Progress -Activity $Activity -Completed

    If ($AsJob.IsPresent) {
        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) submitted; run 'Get-Job x | Wait-Job | Receive-Job'"
            Write-Verbose $Msg
            $Jobs | Get-Job @StdParams
        }
        Else {
            $Msg = "No jobs running"
            $Host.UI.WriteErrorLine($Msg)
        }
    }
    Else {
        If ($Results.Count -gt 0) {
            Write-Output $Results | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceId
        }
        Else {
            $Msg = "No results"
            $Host.UI.WriteErrorLine($Msg)
        }
    }

}
} #end Install-PKWindowsPIPPackagePackage


$Null = New-Alias -Name Install-PKWindowsPipPackage -Value Install-PKWindowsPythonPackage -Force -Description "Consistency and searchability" -Confirm:$False
$Null = New-Alias -Name Install-PKWindowsPip -Value Install-PKWindowsPythonPackage -Force -Description "Consistency and searchability" -Confirm:$False
$Null = New-Alias -Name Install-PKPipPackage -Value Install-PKWindowsPythonPackage -Force -Description "Consistency and searchability" -Confirm:$False