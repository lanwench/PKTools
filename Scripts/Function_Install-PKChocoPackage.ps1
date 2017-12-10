#requires -Version 3
Function Install-PKChocoPackage {
<#
.SYNOPSIS
    Uses Invoke-Command to install Chocolatey packages on one or more computers, interactively or as PSJobs

.DESCRIPTION
    Uses Invoke-Command to install Chocolatey packages on one or more computers, interactively or as PSJobs
    Accepts pipeline input
    Returns a PSJob object
    

.NOTES
    Name    : Function_Install-PKChocoPackage.ps1
    Created : 2017-06-27
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :

        # PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v1.0.0      - 2017-06-27 - Created script
        v01.01.0000 - 2017-12-01 - Updates, made job optional


.EXAMPLE
    PS C:\> Install-PKChocoPackage -ComputerName ops-pktest-1 -Package python -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        ComputerName          {ops-pktest-1}                           
        Package               {python}                                 
        Verbose               True                                     
        Force                 False                                    
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        SuppressconsoleOutput False                                    
        ScriptName            Install-PKChocoPackage                   
        ScriptVersion         1.1.0                                    

        Action: Install Chocolatey package
        VERBOSE: ops-pktest-1


        ComputerName      : OPS-PKTEST-1
        Package           : python
        Version           : (latest)
        InstalledVersion  : {python 3.6.3, python2 2.7.13, python3 3.6.3}
        AvailableVersion  : 3.6.3 
        InstallSuccessful : False
        ScriptOutput      : -
        Messages          : Latest available version of python already installed; -Force not specified


.EXAMPLE
    PS C:\> Install-PKChocoPackage -ComputerName ops-pktest-1 -Package 'python 9.9.9' -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        ComputerName          {ops-pktest-1}                           
        Package               {python 9.9.9}                           
        Verbose               True                                     
        Force                 False                                    
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        SuppressconsoleOutput False                                    
        ScriptName            Install-PKChocoPackage                   
        ScriptVersion         1.1.0                                    

        Action: Install Chocolatey package
        VERBOSE: ops-pktest-1


        ComputerName      : OPS-PKTEST-1
        Package           : python
        Version           : 9.9.9
        InstalledVersion  : Error
        AvailableVersion  : -
        InstallSuccessful : False
        ScriptOutput      : Error
        Messages          : No match found; verify package name/version



















.EXAMPLE
    PS C:\> Install-PKChocoPackage -ComputerName ops-pktest-1 -Package sysinternals -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key              Value                                    
        ---              -----                                    
        ComputerName     {ops-pktest-1}                           
        Package          sysinternals                             
        Verbose          True                                     
        Credential       System.Management.Automation.PSCredential
        Force            False                                    
        JobName          Choco                                    
        ScriptName       Install-PKChocoPackage                   
        ScriptVersion    1.0.0                                                        

        VERBOSE: ops-pktest-1

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        30     Choco           RemoteJob       Running       True            ops-pktest-1         ...                      


        PS C:\> Get-Job 30 | Receive-Job 

        ComputerName     : OPS-PKTEST-1
        Package          : sysinternals
        AlreadyInstalled : False
        Installed        : True
        Messages         : Package installed successfully
        Output           : {Chocolatey v0.10.3, Installing the following packages:, sysinternals, By installing you accept licenses for the 
                           packages....}
        PSComputerName   : ops-pktest-1
        RunspaceId       : 64263ddd-1156-4c26-8544-53c3e9dd2586

.EXAMPLE
    PS C:\> (Get-ADComputer ops-pktest-1).DNSHostName | 
        Install-PKChocoPackage -Package nodejs.install -Credential $Credential -JobName "MyJob" -Verbose | Get-Job | Wait-Job | Receive-Job
        
        VERBOSE: PSBoundParameters: 
	
        Key           Value                                    
        ---           -----                                    
        Package       nodejs.install                           
        Credential    System.Management.Automation.PSCredential
        JobName       MyJob                                    
        Verbose       True                                     
        ComputerName                                           
        Force         False                                    
        ScriptName    Install-PKChocoPackage                   
        ScriptVersion 1.0.0                                    

        VERBOSE: ops-pktest-1.domain.local

        ComputerName     : OPS-PKTEST-1
        Package          : nodejs.install
        AlreadyInstalled : False
        Installed        : True
        Messages         : Package installed successfully
        Output           : {Chocolatey v0.10.3, Installing the following packages:, nodejs.install, By installing you accept licenses for the 
                           packages....}
        PSComputerName   : ops-pktest-1.domain.local
        RunspaceId       : b38f981a-1044-4fd2-97b7-d2464054526d

#>
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [parameter(
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage = "Name of target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("NodeName","Hostname","Name","FQDN")]
    [string[]]$ComputerName,

    [parameter(
        Mandatory=$True,
        HelpMessage = "Package name, or name and version (e.g., 'python' or 'python 3.6.3' or 'python v2.7.0') "
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$Package,

        [parameter(
        Mandatory=$False,
        HelpMessage = "Force install"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$Force,

    [parameter(
        Mandatory=$False,
        HelpMessage = "Valid credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,

    [parameter(
        ParameterSetName = "Job",
        Mandatory=$False,
        HelpMessage = "Run as PSJob"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$AsJob,

    [parameter(
        Mandatory=$False,
        HelpMessage = "Suppress non-verbose, non-error console output"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$SuppressconsoleOutput

)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.000"

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Scriptblock
    
    $ScriptBlock = {
        Param($Package,$Force)
        
        [string[]]$Package = $Using:Package
        [switch]$Force = $Using:Force

        # Since package param accepts multiple entries, didn't want to create
        # another property for version
        $PackageSplit = @()
        $Package | Foreach-Object {
            
            If ($IncludesVersion = ($_ | Where-Object {$_ -match "\ ?[0-9|.]+"})){
                $Version = ($IncludesVersion -Split("^\w+\ ") | Where-Object {$_}).Trim()
                $Name = ($IncludesVersion -Split(" ?[0-9|\.]+") | Where-Object {$_}).Trim()
                $Display = "$Name $Version"
            }
            Else {
                $Name = $_
                $Version = "(latest)"
                $Display = "Latest available version of $Name"
            }
            $PackageSplit += New-Object PSObject -Property @{
               Name = $Name.Trim()
               Version = $Version
               Display = $Display
            }
        }

        # Output object
        $Results = @()
        $InitialValue = "Error"
        $OutputTemplate = New-Object PSObject -Property @{
            ComputerName      = $Env:ComputerName
            Package           = $PackageSplit.Name
            Version           = $PackageSplit.version
            Force             = $Force
            InstalledVersion  = $InitialValue
            AvailableVersion  = $InitialValue
            InstallSuccessful = $InitialValue
            LogPath           = $InitialValue
            Messages          = $InitialValue
            ScriptOutput      = $InitialValue
        }
        $Select = "ComputerName","Package","Version","Force","InstalledVersion","AvailableVersion","InstallSuccessful","LogPath","ScriptOutput","Messages"

        #Make sure chocolatey is installed 
        [switch]$Continue = $False
        Try {
            If ($Choco = Get-Command choco.exe -ErrorAction Stop) {
                $Continue = $True
                $LogFile = (Get-Item $Choco.Source).Directory.Parent.FullName | Get-Childitem -Include chocolatey.log -Recurse
                $OutputTemplate.LogPath = $LogFile.FullName
            }
            Else {
                $Msg = "Chocolatey not found; see https://chocolatey.org"
                $OutputTemplate.InstallSuccessful = $False
                $OutputTemplate.Messages = $Msg
                $Results += $OutputTemplate
            }
        }
        Catch {
            $Msg = "Chocolatey not found"
            If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$Msg"}
            $OutputTemplate.InstallSuccessful = $False
            $OutputTemplate.Messages = $Msg
            $Results += $OutputTemplate
        }

        # If chocolatey found
        If ($Continue.IsPresent) {
            
            $Continue = $False 

            Foreach ($Pkg in $PackageSplit) {
                
                $Continue = $False

                $Output = $OutputTemplate.PSObject.Copy()
                $Name = $Output.Package = $Pkg.Name
                $Version = $Output.Version = $Pkg.Version

                $RegexStr = "($Name)(\ |v|[0-9]|\||\.)+"
                [regex]$Regex = $RegexStr

                # Will either use choco search or choco upgrade 
                Switch ($Pkg.Version) {
                     Default {
                        [Switch]$GetLatest = $True
                        
                        $SearchCommand = "choco search $Name --version $Version --exact --approved-only --not-broken -r"
                        $ListCommand = "choco list $Name --version $Version --local-only"
                   
                        If ($GetAvailable = (Invoke-Expression -Command $SearchCommand -ErrorAction Stop | Select-String $RegexStr | Where-Object {$_}) -as [string]) {
                            $Output.AvailableVersion = $Version
                            $Continue = $True
                        }
                        Else {
                            $Msg = "No match found; verify package name/version"
                            $Output.AvailableVersion = "-"
                            $Output.InstallSuccessful = $False
                            $Output.Messages = $Msg
                        }
                    }
                    '(latest)' {
                        [Switch]$GetLatest = $False
                        
                        $SearchCommand = "choco upgrade $Name -whatif"
                        $ListCommand = "choco list $Name --local-only"
                        $FoundStr = "is the latest version available"
                        $NotFoundStr = "The package was not found"
                    
                        If ($Search = Invoke-Expression -Command $SearchCommand -ErrorAction Stop) {
                            
                            Switch -regex ($Search) {
                                $FoundStr {
                                    "Found it"
                                    ($Found = ($Search | Select-String $RegexStr | Where-Object {$_}) -as [string])

                                    # Get the name/version info from the upgrade command
                                    $LatestVer = $Regex.Matches($Found) | Foreach-Object {$_.Value}
                                
                                    # Update the display name
                                    $Display = $LatestVer.Trim()
                                    $Output.AvailableVersion = $LatestVer.split("v")[1]
                                
                                    $Continue = $True
                                }
                                $NotFoundStr {

                                    "nope"
                                    $Msg = "No matching package found"
                                    $Output.AvailableVersion = "-"
                                    $Output.InstallSuccessful = $False
                                    $Output.Messages = $Msg
                                }
                            }
                        }
                        Else {
                            $Msg = "Package search failed"
                            $Output.AvailableVersion = "-"
                            $Output.InstallSuccessful = $False
                            $Output.Messages = $Msg
                        }
                            <#
                            
                            
                            
                            
                            
                            
                            
                            
                            }

                            If ($Found = ($Search | Select-String $RegexStr | Where-Object {$_}) -as [string]) {
                                
                                If ($Found -match "The package was not found") {
                                    $Outp


                                }

                                Else {
                                    # Get the name/version info from the upgrade command
                                    $LatestVer = $Regex.Matches($Found) | Foreach-Object {$_.Value}
                                
                                    # Update the display name
                                    $Display = $LatestVer.Trim()
                                    $Output.AvailableVersion = $LatestVer.split("v")[1]
                                
                                    $Continue = $True
                                }
                                
                            }
                            Elseif (($Search | Select-String $NotFoundStr | Where-Object {$_}) -as [string]) {
                                $Msg = "No matching package found"
                                $Output.AvailableVersion = "-"
                                $Output.InstallSuccessful = $False
                                $Output.Messages = $Msg
                            }
                        }
                    }
                } #end switch

                #>
                    } #end if latest version
                }

                # Look for current install
                If ($Continue.IsPresent) {
                    
                    $Continue = $False

                    Try {
                        If ($GetInstalled = (Invoke-Expression -Command $ListCommand -ErrorAction Stop  | Select-String $RegexStr | Where-Object {$_})) {
                        
                            # List all installed / matches
                            $Output.InstalledVersion= ($GetInstalled | Where-Object {$_})

                            # See if we already have it
                            If (($GetInstalled | Where-Object {$_}) | Select-String $Regex) {
                                If (-not $Force.IsPresent) {
                                    $Msg = "$Display already installed; -Force not specified"
                                    $Output.Messages = $Msg
                                    $Output.InstallSuccessful = $False
                                    $Output.ScriptOutput = "-"
                                }
                                Else {
                                    $Continue = $True
                                }
                            }
                            Else {
                                $Continue = $True
                            }
                        }
                    }
                    Catch {
                        $Msg = "Package installation test failed"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                        $Output.InstallSuccessful = $False
                        $Output.Messages = $Msg
                
                    }
                } #end if continue

                # Try the installation
                If ($continue.IsPresent) {
                    Try {
                        [string]$InstallCommand = "choco install $Name -y"
                        If ($Force.IsPresent) {
                            $InstallCommand +=" --force"
                        }
                        
                        $Install = Invoke-Expression -Command $InstallCommand -ErrorAction Stop
                        $Output.ScriptOutput = $Install

                        Switch -Regex ($Install) {
                            "The install of $Name was successful" {
                                $Output.InstallSuccessful = $True
                                $Output.Messages = "Package installed successfully; see logfile"
                            }
                            "The package was not found" {
                                $Output.InstallSuccessful = $False
                                $Output.Messages = "Package not found"
                            }
                            "Package failed to install" {
                                $Output.InstallSuccessful = $False
                                $Output.Messages = "Package failed to install; see logfile"
                            }
                            "The install of $Name was not successful" {
                                $Output.InstallSuccessful = $False
                                $Output.Messages = "Package failed to install; see logfile"
                            }
                            Default {
                                $Output.InstallSuccessful = $False
                                $Output.Messages = "Package failed to install; see logfile"
                            }
                        }
                    }
                    Catch {
                        $Msg = "Installation failed"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$Msg"}
                        $Output.InstallSuccessful = $False
                        $Output.Messages = $Msg
                    }
                }

                $Results += $Output

            } # end for each package
       
       } #end if continue
                
       Write-Output ($Results | Select $Select)

    } #end scriptblock

    #endregion Scriptblock

    #region Splats

    # Activity
    If ($SkipConnectionTest.IsPresent) {[string]$Activity = "Test connection and install Chocolatey "}
    Else {[string]$Activity = "Install Chocolatey "}
    If ($Package.Count -gt 1) {
        $PkgStr = "'$($Package -join("', '"))'"
        $Activity += "packages $PkgStr"
    }
    Else {
        $Activity += "package $Package"
    }
    If ($Force.IsPresent) {$Activity += " (force)"}


    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for write-progress
    #$Activity = "Install Chocolatey package"
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
        ArgumentList = $Package,$Force
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

    # Output object if not running as job
    $Results = @()

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
                
        [Switch] $Continue = $False
        If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
            
            If (-not $SkipConnectionTest.IsPresent) {
                
                If ($Computer -eq $env:COMPUTERNAME) {
                    $Msg = "(Skipping connection test to local computer)"
                    Write-Verbose $Msg
                    $Continue = $True
                }
                Else {                        
                    Try {
                        $Param_WSMAN.ComputerName = $Computer
                        If ($Null = Test-WSMan @Param_WSMan ) {
                            $Continue = $True
                        }
                    }
                    Catch {
                        $Msg = "Connection failure on $Computer"
                        $ErrorDetails= [regex]:: match($_.Exception.Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()
                        $Host.UI.WriteErrorLine("ERROR: $Msg [$ErrorDetails]")
                    }
                }
            } #end if testing WinRM
            Else {$Continue = $True}
            
            If ($Continue.IsPresent) {
                
                $Param_IC.ComputerName = $Computer

                If ($AsJob.IsPresent) {
                    $Param_IC.JobName = "Choco_$Computer"
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

} #end Install-PKChocoPackage

