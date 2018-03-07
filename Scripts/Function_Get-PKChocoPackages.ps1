#Requires -version 3
Function Get-PKChocoPackages {
<#
.SYNOPSIS
    Gets a list of locally installed Chocolatey packages, interactively or as a PSJob

.DESCRIPTION
    Gets a list of locally installed Chocolatey packages, interactively or as a PSJob
    Optionally returns packages/versions as a collection
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES
    Name    : Function_Get-PKChocoPackages
    Created : 2018-01-29
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2018-01-29 - Created script

.PARAMETER ComputerName
    Name of computer to do cool thing on; separate multiple names with commas

.PARAMETER ExpandPackages
    Return package names/versions as a collection

.PARAMETER Credential
    Valid credentials on target computer (default is current user credentials)

.PARAMETER AsJob
    Run as PSJob

.PARAMETER SkipConnectionTest
    Don't test WinRM connectivity before submitting command

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKChocoPackages -Verbose

        VERBOSE: PSBoundParameters:

       Key                   Value
        ---                   -----
        Verbose               True
        ComputerName
        ExpandPackages        False
        Credential            System.Management.Automation.PSCredential
        AsJob                 False
        SkipConnectionTest    False
        SuppressConsoleOutput False
        PipelineInput         False
        ScriptName            Get-PKChocoPackages
        ScriptVersion         1.0.0


        ComputerName : PAULA-WS
        ChocoVersion : 0.10.5.0
        NumPackages  : 32
        Packages     : {7zip 16.4.0.20170506, 7zip.install 16.4.0.20170506, chocolatey 0.10.8, chocolatey-core.extension 1.3.1...}
        Messages     : 32 package(s) found


.EXAMPLE
    PS C:\> $Arr | Get-PKChocoPackages -Verbose

        VERBOSE: PSBoundParameters:

        Key                   Value
        ---                   -----
        Verbose               True
        ComputerName
        ExpandPackages        False
        Credential            System.Management.Automation.PSCredential
        AsJob                 False
        SkipConnectionTest    False
        SuppressConsoleOutput False
        PipelineInput         True
        ScriptName            Get-PKChocoPackages
        ScriptVersion         1.0.0

        Action: List locally-installed Chocolatey packages
        VERBOSE: TEST-VM-1
        VERBOSE: DEV-APP12
        VERBOSE: foo.domain.local
        ERROR: WinRM connection failed on foo.domain.local

        ComputerName : TEST-VM-1
        ChocoVersion : 0.10.5.0
        NumPackages  : 13
        Packages     : {chocolatey 0.10.8, chocolatey-core.extension 1.3.0, easy.install 0.6.11.4, GoogleChrome 58.0.3029.96...}
        Messages     : 13 package(s) found

        ComputerName : DEV-APP12
        ChocoVersion : 0.9.10.0
        NumPackages  : 6
        Packages     : {chocolatey 0.10.3, jdk8 8.0.131, python 3.6.3, python2 2.7.13...}
        Messages     : 6 package(s) found


.EXAMPLE
    PS C:\> $Computers | Get-PKChocoPackages -ExpandPackages -AsJob -SkipConnectionTest -SuppressConsoleOutput

        Id     Name           PSJobTypeName   State      HasMoreData   Location               Command
        --     ----           -------------   -----      -----------   --------               -------
        51     Choco_SQLN...  RemoteJob       Completed  True          SQLNEW                 ...
        53     Choco_labm...  RemoteJob       Running    True          labmachine             ...
        55     Choco_devb...  RemoteJob       Running    True          devbox                 ...
        57     Choco_lega...  RemoteJob       Failed     False         legacyapp.domain.local ...


        [...]

        PS C:\> Get-Job choco* | Where-Object {$_.State -eq "Completed"}

        ComputerName : SQLNEW
        ChocoVersion : Error
        NumPackages  : Error
        Packages     : Error
        Messages     : Chocolatey not found; see https://chocolatey.org

        ComputerName : LABMACHINE
        ChocoVersion : 0.10.5.0
        NumPackages  : 13
        Packages     : @{sysinternals=2017.6.14; jdk8=8.0.131; easy.install=0.6.11.4; python=3.6.3; python3=3.6.3; chocolatey=0.10.8;
                       GoogleChrome=58.0.3029.96; sqlserver-odbcdriver=13.1.4413.46; pip=1.2.0; python2=2.7.13; chocolatey-core.extension=1.3.0;
                       sqlserver-cmdlineutils=13.1; nodejs.install=8.1.2}
        Messages     : 13 package(s) found

        ComputerName : DEVBOX
        ChocoVersion : 0.9.10.0
        NumPackages  : 6
        Packages     : @{jdk8=8.0.131; python3=3.6.3; chocolatey=0.10.3; python=3.6.3; vcredist2013=12.0.30501.20150616; python2=2.7.13}
        Messages     : 6 package(s) found


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
        Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Hostname or FQDN of computer (separate multiple computers with commas)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory = $False,
        HelpMessage = "Expand package names/versions as a collection"
    )]
    [Switch] $ExpandPackages,

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


    # Detect pipeline input & save parametersetname
    $Source = $PSCmdlet.ParameterSetName
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and ((-not $ComputerName))

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} |
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    #$CurrentParams.Add("ParameterSetName",$Source)
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

    # If we didn't supply anything
    If (-not $ComputerName) {$ComputerName = $Env:ComputerName}

    # Output
    $Results = @()

    # Scriptblock for invoke-command
    $ScriptBlock = {
        Param([switch]$ExpandPackages)
        $ErrorActionPreference = "Stop"

        $InitialValue = "Error"
        $Output = New-Object PSObject -Property @{
            ComputerName = $Env:ComputerName
            ChocoVersion = $InitialValue
            NumPackages  = $InitialValue
            Packages     = $InitialValue
            Messages     = $InitialValue
        }

        $Select = "ComputerName","ChocoVersion","NumPackages","Packages","Messages"

        Try {
            If ($Choco = Get-Command choco.exe -ErrorAction SilentlyContinue) {

                $Output.ChocoVersion = $Choco.Version.ToString()

                $Cmd = "choco list -localonly"

                [array]$AllPackages = Invoke-Expression -Command $Cmd -EA Stop

                $Output.ChocoVersion = $Choco.Version.ToString()

                [array]$FoundPackages = ((($AllPackages -split($AllPackages[-1])) -split($AllPackages[0])) | Where-Object {$_})
                $Output.NumPackages = $FoundPackages.Count

                If ($FoundPackages.Count -gt 0) {
                    $Msg = "$($FoundPackages.Count) package(s) found"
                    $Output.Messages = $Msg

                    If ($Using:ExpandPackages.IsPresent){
                        $Expanded = @{}
                        $FoundPackages | Foreach-Object {
                            $Expanded.Add($_.Split(" ")[0],($_.Split(" ")[1]))
                        }
                        $Output.Packages = New-Object -TypeName PSObject -Property $Expanded
                    }
                    Else {
                        $Output.Packages = $FoundPackages
                    }
                }
                Else {
                    $Msg = "No packages found"
                    $Output.Packages = "-"
                    $Output.Messages = $Msg
                }
            }
            Else {
                $Msg = "Chocolatey not found; see https://chocolatey.org"
                $Output.Messages = $Msg
            }
        }
        Catch {
            $Msg = "Chocolatey not found; see https://chocolatey.org"
            If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
            $Output.Messages = $Msg
        }

        Write-Output ($Output | Select-Object $Select)

    } #end scriptblock

    #region Splats

    # Splat for Write-Progress
    $Activity = "List locally-installed Chocolatey packages"
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
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $ExpandPackages
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.AsJob = $True
        $Param_IC.JobName = $Null
        $JobPrefix = "Choco"
    }

    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}


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

        $Param_WP.Status = $Computer
        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.PercentComplete = $PercentComplete

        [switch]$Continue = $False

        # If we're testing WinRM
        If (-not $SkipConnectionTest.IsPresent -and ($Computer -notin @("$Env:ComputerName","Localhost","127.0.0.1"))) {

            $Msg = "Test WinRM connection"
            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP

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

                    $Msg = "Submit PSJob"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC
                        $Msg = "Job ID $($Job.ID): $($Job.Name)"
                        Write-Verbose $Msg
                        $Jobs += $Job
                    }
                    Else {
                        $Msg = "Invoke scriptblock"
                        Write-Progress @Param_WP
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
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            Write-Verbose $Msg
            $Jobs | Get-Job

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

} # end Get-PKChocoPackages

$Null = New-Alias -Name Get-PKWindowsChocoPackages -Value Get-PKChocoPackages -Force -ErrorAction SilentlyContinue -Verbose:$False -Description "2018-01-29 consistency" -Confirm:$False