#requires -Version 3
Function Get-PKWindowsPath {
<#
.Synopsis
    Gets the environment path value for a computer or user or both

.DESCRIPTION
    Gets the environment path value for a computer or user or both
    Uses Invoke-Command
    Accepts pipeline input
    Returns a PSObject

.NOTES
    Name    : Function_Get-PKWindowsPath.ps1
    Created : 2017-10-29
    Version : 01.00.0000
    Author  : Paula Kingsley
    History:
        
        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v01.00.0000 - 2017-10-29 - Created script


.PARAMETER ComputerName
    Computer name. Defaults to local computer. 

.PARAMETER Context
    Path for user, machine, or all. For remote computers, user will likely return null.

.PARAMETER Unique
    Filter out duplicate path entries (per context)

.PARAMETER StackOutput
    If a collection is returned, display as a multiline string

.PARAMETER Credential
    Valid credential on target computer (default is current user)

.PARAMETER AsJob
    Run as PowerShell job and return PSJob objects

.PARAMETER SkipConnectionTest
    Don't test WinRM connection before running Invoke-Command

.PARAMETER SuppressConsoleOutput
    Suppress non-verbose console output

.EXAMPLE
    PS C:\> Get-PKWindowsPath -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Verbose               True                                     
        ComputerName          {WORKSTATION}                        
        Context               All                                      
        Unique                False                                    
        StackOutput           False                                    
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        SkipConnectionTest    False                                    
        SuppressConsoleOutput False                                    
        PipelineInput         False                                    
        ScriptName            Get-PKWindowsPath                        
        ScriptVersion         1.0.0                                    

        Action: Test connection and get environment path for both machine and user
        VERBOSE: WORKSTATION
        VERBOSE: Performing the operation "Test connection and get environment path for both machine and user" on target "WORKSTATION".
        VERBOSE: (Skipping connection test to local computer)

        ComputerName    Context Path                                                                                                                 
        ------------    ------- ----                                                                                                                 
        WORKSTATION Machine {C:\Distrib\, C:\HashiCorp\Vagrant\bin, C:\opscode\chefdk\bin\, C:\opscode\chefdk\embedded...}                       
        WORKSTATION User    {C:\Program Files\Microsoft VS Code\bin, C:\Program Files\Oracle\VirtualBox, C:\Users\pkingsley\AppData\Local\Pand...

.EXAMPLE
    PS C:\> $Arr | Get-PKWindowsPath -Context Machine -Verbose | Format-List

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Context               Machine                                  
        Verbose               True                                     
        ComputerName          {}                        
        Unique                False                                    
        StackOutput           False                                    
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        SkipConnectionTest    False                                    
        SuppressConsoleOutput False                                    
        PipelineInput         False                                    
        ScriptName            Get-PKWindowsPath                        
        ScriptVersion         1.0.0                                    


        Action: Test connection and get environment path for machine
        VERBOSE: SQLSERVER-2
        VERBOSE: Performing the operation "Test connection and get environment path for machine" on target "SQLSERVER-2".
        VERBOSE: foo
        VERBOSE: Performing the operation "Test connection and get environment path for machine" on target "foo".
        ERROR: Connection failure on foo (The WinRM client cannot process the request because the server name cannot be resolved.)
        VERBOSE: webtest
        VERBOSE: Performing the operation "Test connection and get environment path for machine" on target "webtest".


        ComputerName : SQLSERVER-2
        Context      : Machine
        Path         : {C:\AP\apache-maven-3.3.9\bin, C:\opscode\chef\bin\, C:\opscode\chef\embedded\bin\, C:\oracle\product\11.2.0\client_32\bin...}

        ComputerName : WEBTEST
        Context      : Machine
        Path         : {C:\opscode\chef\bin\, C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\130\Tools\Binn\, C:\ProgramData\chocolatey\bin, 
                       C:\Python36\...}


.EXAMPLE
    PS C:\> Get-PKWindowsPath -ComputerName server99 -Context Machine -Unique -StackOutput -Credential $Credential -AsJob -SkipConnectionTest -SuppressConsoleOutput 

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        1      Path_auto2-d... RemoteJob       Completed     False           server99          ...         

        [...]

        PS C:\> Get-Job 1 | Receive-Job

        ComputerName   : server99
        Context        : Machine
        Path           : C:\app\oracle\product\11.2.0\client_1\bin
                         C:\opscode\chef\bin\
                         C:\Program Files\Electric Cloud\ElectricCommander\bin
                         C:\Program Files\Java\jdk1.8.0_131\bin
                         c:\Program Files\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn\
                         C:\Program Files\Perforce
                         C:\ProgramData\chocolatey\bin
                         C:\Windows
                         C:\Windows\system32
                         C:\Windows\System32\Wbem
                         C:\Windows\System32\WindowsPowerShell\v1.0\
        PSComputerName : server99
        RunspaceId     : a4745ae1-3b7f-4c4b-b1c7-a1f74f4a99bc

#>
[Cmdletbinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "Low"
)]
Param(

    [Parameter(
        Mandatory = $False,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Name of computer (separate multiple computers with commas)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name","HostName","FQDN","DNSDomainName")]
    [String[]] $ComputerName = $Env:ComputerName,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Path type (user, machine, or all). Note that for remote computers, 'user' will likely be null"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("User","Machine","All")]
    [String] $Context = "All",

    [Parameter(
        Mandatory=$false,
        HelpMessage="Return unique paths only (per source/context)"
    )]
    [Switch] $Unique,

    [Parameter(
        Mandatory=$false,
        HelpMessage="Return path output as a stacked string (individual lines) instead of a collection"
    )]
    [Switch] $StackOutput,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Valid credentials on target computer(s)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty ,

    [Parameter(
        Mandatory=$false,
        HelpMessage="Run as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        Mandatory=$false,
        HelpMessage="Don't test WinRM connection before submitting command"
    )]
    [Switch] $SkipConnectionTest,        

    [Parameter(
        Mandatory=$false,
        HelpMessage="Suppress non-verbose console output"
    )]
    [Switch] $SuppressConsoleOutput
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    #$Source = $PSCmdlet.ParameterSetName
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputeName")) -and (-not $ComputerName)

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

    # Create scriptblock
    $Scriptblock = {
        
        Param($Context,$StackOutput,$Unique)
        $Type = $Using:Context
        If ($Using:StackOutput) {[switch]$Stack = $True}
        If ($Using:Unique) {[switch]$U = $True}

        $StdParams = @{}
        $StdParams = @{
            ErrorAction = "Stop"
            Verbose     = $False
        }

        $OutputHT = @{
            ComputerName = $Env:ComputerName
            Context      = "Error"
            Path         = "Error"
            Messages     = "Error"
        }
        $Select = "ComputerName","Context","Path"

        $Results = @()

        Try {
            Switch ($Type) {
                User {
                    $Output = New-Object PSObject -Property $OutputHT
                    $Output.Context = "User"
                    $Path = ([environment]::GetEnvironmentVariable("Path","User") -split(";") | Where-Object {$_} | Sort-Object)
                    If ($U.IsPresent) {$Path = $Path | Select-Object -Unique}
                    If ($Stack.IsPresent) {$Path = $Path -join("`n")}
                    $Output.Path = $Path 
                    $Output.Messages = $Null
                    $Results = $Output 
                }
                Machine {
                    $Output = New-Object PSObject -Property $OutputHT
                    $Output.Context = "Machine"
                    $Path = [environment]::GetEnvironmentVariable("Path","Machine") -split(";") | Where-Object {$_} | Sort-Object
                    If ($U.IsPresent) {$Path = $Path | Select-Object -Unique}
                    If ($Stack.IsPresent) {$Path = $Path -join("`n")}
                    $Output.Path = $Path 
                    $Output.Messages = $Null
                    $Results = $Output 
                }
                All {
                    $Output = New-Object PSObject -Property $OutputHT
                    $Path = [environment]::GetEnvironmentVariable("Path","Machine") -split(";") | Where-Object {$_} | Sort-Object
                    If ($U.IsPresent) {$Path = $Path | Select-Object -Unique}
                    If ($Stack.IsPresent) {$Path = $Path -join("`n")}
                    $Output.Context = "Machine"
                    $Output.Path = $Path
                    $Output.Messages = $Null
                    $Results += $Output 

                    $Output = New-Object PSObject -Property $OutputHT
                    $Path = [environment]::GetEnvironmentVariable("Path","User") -split(";") | Where-Object {$_} | Sort-Object
                    If ($U.IsPresent) {$Path = $Path | Select-Object -Unique}
                    If ($Stack.IsPresent) {$Path = $Path -join("`n")}
                    $Output.Context = "User"
                    $Output.Path = $Path
                    $Output.Messages = $Null
                    $Results += $Output 
                }
            }
        }
        Catch {
            $Output = New-Object PSObject -Property $OutputHT
            $Output.Messages = $_.Exception.Message
            $Results += $Output 
            $Select = "ComputerName","Context","Path","Messages"
        }

        Write-Output $Results | Select-Object $Select

    } # end scriptblock

    Switch ($Context) {
        User {$TypeStr = "user"}
        Machine {$TypeStr = "machine"}
        All {$TypeStr = "both machine and user"}
    }
    
    If ($SkipConnectionTest.IsPresent) {$Activity = "Get environment path for $Typestr"}
    Else {$Activity = "Test connection and get environment path for $Typestr"}
    If ($AsJob.IsPresent) {$Activity = $Activity + " as PSJob"}
    
    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for test-wsman
    $Param_WSMan = @{}
    $Param_WSMan = @{
        ComputerName   = $Null
        Authentication = "Negotiate"
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
        ArgumentList = $OutputType,$StackOutput
        ErrorAction  = "Stop"
        Verbose      = $False
    }
    If ($AsJob.IsPresent) {
        $Param_IC.Add("AsJob",$True)
        $Param_IC.Add("JobName",$Null)
        $Jobs = @()
    }
    
    # Splat for write-progress
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        PercentComplete  = $Null
        CurrentOperation = $Null
        Status           = "Working"
    }

    #endregion Splats

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
                
        [Switch] $Continue = $True
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
                        If ($Null = Test-WSMan @Param_WSMan ) {$Continue = $True}
                    }
                    Catch {
                        $Msg = "Connection failure on $Computer"
                        $ErrorDetails= [regex]:: match($_.Exception.Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()
                        $Host.UI.WriteErrorLine("ERROR: $Msg ($ErrorDetails)")
                        $Continue = $False
                    }
                }
            } #end if testing WinRM
            Else {$Continue = $True}
            
            If ($Continue.IsPresent) {
                
                $Param_IC.ComputerName = $Computer
                If ($AsJob.IsPresent) {
                    $Param_IC.JobName = "Path_$Computer"
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
        Write-Output $Results | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID
    }

}
} #end Get-PKWindowsPath

