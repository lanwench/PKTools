#Requires -version 3
Function Get-PKTaniumActionLog {
<# 
.SYNOPSIS
    Invokes a scriptblock to parse content from the Tanium client action log, returning a PSObject

.DESCRIPTION
    Invokes a scriptblock to parse content from the Tanium client action log, returning a PSObject
    Allows for filtering on action type (via regex) and by start/end date
    Allows for selection of only most recent log file
    Accepts pipeline input
    Optionally tests connectivity to remote computers before invoking scriptblock
    Returns a PSobject or PSJob

.NOTES        
    Name    : Get-PKTaniumActionLog.ps1
    Created : 2019-06-19
    Author  : Paula Kingsley
    Version : 02.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-06-19 - Created script
        v02.00.0000 - 2019-08-13 - Added parameters for date, action type, handling for multiple action files

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER CustomFilePath
    Absolute path to Tanium client action logfile (default is to look for action-history text file in c:\Program Files (x86)\Tanium\Tanium Client\Logs)

.PARAMETER AllFiles
    Include all discovered action-history log files (default is the most recently written to)

.PARAMETER ActionTypeFilter
    Return only actions where the filter string matches an action type (default is no filter)

.PARAMETER StartDate
    Start date (default is earliest in logs)

.PARAMETER EndDate
    End date (default is now)

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
	
        Key              Value                                    
        ---              -----                                    
        Verbose          True                                     
        ComputerName     LAPTOP                          
        CustomFilePath                                            
        AllFiles         False                                    
        ActionTypeFilter                                          
        StartDate                                                 
        EndDate          2019-08-13 04:01:13 PM                   
        Credential       System.Management.Automation.PSCredential
        AsJob            False                                    
        JobPrefix        TaniumActivity                           
        ConnectionTest   WinRM                                    
        NoProgress       False                                    
        Quiet            False                                    
        ScriptName       Get-PKTaniumActionLog                    
        ScriptVersion    1.1.0                                    
        PipelineInput    False                                    

        BEGIN  : Get content of latest Tanium Client action log

        [LAPTOP] Invoke command

        Computername : LAPTOP
        Date         : 2019-06-27 12:22:38 AM
        ActionID     : 1102636
        Action       : Patch - Maintenance Window 38
        Message      : completed in 0.118 seconds with exit code 0

        Computername : LAPTOP
        Date         : 2019-06-27 12:24:48 AM
        ActionID     : 1102644
        Action       : Patch - Scan Configuration 2
        Message      : executing: cmd /c xcopy /y scan-configuration-2.xml ..\..\Patch\scans\configurations\

        Computername : LAPTOP
        Date         : 2019-06-27 12:25:42 AM
        ActionID     : 1102652
        Action       : Patch - Blacklist 3
        Message      : completed in 0.113 seconds with exit code 0

        Computername : LAPTOP
        Date         : 2019-06-27 12:25:58 AM
        ActionID     : 1102653
        Action       : Deploy - Profile 2
        Message      : executing: cmd /c xcopy /y profile_2.json ..\..\Tools\Deploy\profiles\

        Computername : LAPTOP
        Date         : 2019-06-27 12:27:38 AM
        ActionID     : 1102659
        Action       : Patch - Patch Lists - Windows
        Message      : completed in 0.132 seconds with exit code 0

        <snip>

        END    : Get content of latest Tanium Client action log


.EXAMPLE
    PS C:\> Get-PKTaniumActionLog -ComputerName Test-VM.domain.local -Credential $Credential -StartDate 2019-08-01 -EndDate (get-Date).AddDays(-1) -AsJob -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                                    
        ---                      -----                                    
        StartDate                2019-08-01                               
        EndDate                  2019-08-12 02:48:02 PM                   
        AsJob                    True                                     
        Verbose                  True                                     
        ComputerName             Test-VM                          
        CustomFilePath                                                    
        AllFiles False                                    
        Credential               System.Management.Automation.PSCredential
        JobPrefix                TaniumActivity                           
        ConnectionTest           WinRM                                    
        NoProgress               False                                    
        Quiet                    False                                    
        ScriptName               Get-PKTaniumActionLog                    
        ScriptVersion            1.1.0                                    
        PipelineInput            False                                    

        BEGIN  : Get content of latest Tanium Client action log between 2019-08-01 12:00:00 AM and 2019-08-12 02:48:02 PM (as job)

        [Test-VM] Test WinRM connection
        [Test-VM] Invoke command as PSJob

        1 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        6      TaniumActivi... RemoteJob       Running       True            Test-VM      ...                      

        END    : Get content of latest Tanium Client action log between 2019-08-01 12:00:00 AM and 2019-08-12 02:48:02 PM (as job)

        PS C:\> Get-Job 6 | Wait-Job | Receive-Job | Select -First 4

        Computername   : Test-VM
        Date           : 2019-08-01 12:02:01 AM
        ActionID       : 1333771
        Action         : Detect Intel for Windows Revision 718 Sync
        Message        : executing: cmd /c cscript /nologo run-add-intel-package.vbs 2>&1
        PSComputerName : Test-VM
        RunspaceId     : f18779a8-7cc3-4754-8ae7-3d1ce58c7d23

        Computername   : Test-VM
        Date           : 2019-08-01 12:02:02 AM
        ActionID       : 1333771
        Action         : Detect Intel for Windows Revision 718 Sync
        Message        : completed in 0.407 seconds with exit code 0
        PSComputerName : Test-VM
        RunspaceId     : f18779a8-7cc3-4754-8ae7-3d1ce58c7d23

        Computername   : Test-VM
        Date           : 2019-08-01 01:10:20 AM
        ActionID       : 1333799
        Action         : Patch - Maintenance Window 1
        Message        : executing: cmd /c xcopy /y maintenance-window-1.xml ..\..\Patch\maintenance-windows\configurations\
        PSComputerName : Test-VM
        RunspaceId     : f18779a8-7cc3-4754-8ae7-3d1ce58c7d23

        Computername   : Test-VM
        Date           : 2019-08-01 01:10:20 AM
        ActionID       : 1333799
        Action         : Patch - Maintenance Window 1
        Message        : completed in 0.218 seconds with exit code 0
        PSComputerName : Test-VM
        RunspaceId     : f18779a8-7cc3-4754-8ae7-3d1ce58c7d23

.EXAMPLE    
    PS C:\> Get-PKTaniumActionLog -AllFiles -StartDate 2019-05-02 -EndDate 2019-05-10 -ActionType Patch -AsJob -JobPrefix Foo -Quiet

        Id     Name            PSJobTypeName   State         HasMoreData     Location        Command                  
        --     ----            -------------   -----         -----------     --------        -------                  
        27     Foo-PROD-SQL... RemoteJob       Running       True            PROD-SQL-A      ...                      

        PS C:\> Get-Job 27 | Wait-Job | Receive-Job | Select -First 10 | Format-Table -AutoSize

        Computername    Date                   ActionID Action                        Message                                                        
        ------------    ----                   -------- ------                        -------                                                        
        PROD-SQL-A      2019-07-31 10:48:01 AM 1329750  Patch - Maintenance Window 1  executing: cmd /c xcopy /y maintenance-window-1.xml ..\..\Pa...
        PROD-SQL-A      2019-07-31 10:48:02 AM 1329750  Patch - Maintenance Window 1  completed in 1.015 seconds with exit code 0...                 
        PROD-SQL-A      2019-07-31 10:48:02 AM 1329791  Patch - Maintenance Window 35 executing: cmd /c xcopy /y maintenance-window-35.xml ..\..\P...
        PROD-SQL-A      2019-07-31 10:48:02 AM 1329791  Patch - Maintenance Window 35 completed in 0.334 seconds with exit code 0...                 
        PROD-SQL-A      2019-07-31 10:48:02 AM 1329794  Patch - Blacklist 1           executing: cmd /c xcopy /y blacklist-1.xml ..\..\Patch\black...
        PROD-SQL-A      2019-07-31 10:48:02 AM 1329794  Patch - Blacklist 1           completed in 0.342 seconds with exit code 0...                 
        PROD-SQL-A      2019-07-31 10:48:02 AM 1329824  Patch - Maintenance Window 38 executing: cmd /c xcopy /y maintenance-window-38.xml ..\..\P...
        PROD-SQL-A      2019-07-31 10:48:03 AM 1329824  Patch - Maintenance Window 38 completed in 0.285 seconds with exit code 0...                 
        PROD-SQL-A      2019-07-31 10:48:03 AM 1329834  Patch - Scan Configuration 2  executing: cmd /c xcopy /y scan-configuration-2.xml ..\..\Pa...
        PROD-SQL-A      2019-07-31 10:48:03 AM 1329834  Patch - Scan Configuration 2  completed in 0.424 seconds with exit code 0...                 

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
        HelpMessage = "Include all discovered action-history log files (default is the most recently written to)"
    )]
    [Switch] $AllFiles,

    [Parameter(
        HelpMessage = "Return only actions where the filter string matches an action type (default is no filter)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Asset","Clean","Collect","Comply","Delete","Deploy","Detect","End-User","Execute","Patch","Run","UserGroupTagging","Trace","Windows")]
    [string] $ActionTypeFilter,

    [Parameter(
        HelpMessage = "Start date (default is earliest in logs)"
    )]
    #[ValidateNotNullOrEmpty()]
    $StartDate,

    [Parameter(
        HelpMessage = "End date (default is now)"
    )]
    #[ValidateNotNullOrEmpty()]
    $EndDate = [System.DateTime]::Now,

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
    [version]$Version = "02.00.0000"

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

    #region Preferences 
    
    $ErrorActionPreference = "Stop"
    Switch ($NoProgress) {
        $True  {$ProgressPreference = "SilentlyContinue"}
        $False {$ProgressPreference = "Continue"}
    }

    #endregion Preferences
    
    #region Parameter things

    If ($StartDate -and $EndDate) {
        If (-not ($StartDate -as [datetime]) -or -not ($EndDate -as [datetime])) {
            $Msg = "Please enter a valid start date"
            Throw $Msg
        }
        Else {
            $Start = ($StartDate -as [datetime])
            $End = ($EndDate -as [datetime])
            If (-not ($End -ge $Start)) {
                $Msg = "Time travel is not permitted; please enter an end date later than or equal to $(($StartDate -as [datetime]).ToString())"
                Throw $Msg
            }
        }
    }
    Else {$StartDate = $EndDate = ""}

    #endregion Parameter things

    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        
        Param($ArgList)
        If ($ArgList.CustomFilePath) {
            If (Test-Path $CustomFilePath -PathType Container -EA SilentlyContinue) {

                Try {
                    $ActionFile = Get-ChildItem $CustomFilePath -Recurse -Filter action-history*.txt -EA SilentlyContinue | Get-Item
                }
                Catch {
                    $Msg = "No Tanium client activity logfile found in path '$($ArgList.CustomFilePath)'"
                }
            }
            Else {
                $Msg = "Invalid path '$CustomFilePath"
            }
        }
        Else {
                
            $TaniumCommandPath = "$($Env:SystemDrive)\Program Files*\Tanium\Tanium Client\taniumclient.exe"

            Try {
                If ($Parent = Get-Command -Name $TaniumCommandPath | Select-Object $_.Path | Split-Path -Parent) {
                    $ActionFile = Get-ChildItem $Parent -Recurse -Filter action-history*.txt -EA SilentlyContinue | Get-Item
                }
                Else {
                    $Msg = "No Tanium Client action log file found in $($TaniumCommandPath | Select-Object $_.Path | Split-Path -Parent)"
                }
            }
            Catch {
                $Msg = "Something horrible has happened!"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            }
        }
        
        If ($ActionFile) {
            
            # If there's more than one, and we didn't want all
            If ($ActionFile.Count -gt 1 -and -not $AllFiles.IsPresent) {
                $ActionFile = $ActionFile | Sort LastWriteTime -Descending | Select -First 1
            }
            Try { 

                If ($LogContent = ($ActionFile | Get-Content -Raw -ErrorAction SilentlyContinue)) {       
                    
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
                        $Msg = "No match found in $($ActionFile.FullName -join(", "))"
                    }
                }
                Else {
                    $Msg = "Failed to get content in $($ActionFile.FullName -join(", "))"
                }
            } #end try getting content
            Catch {
                $Msg = "Something horrible has happened!"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            }
        } #end if input file
            
        If (-not $Output) {
            $Output = [pscustomobject]@{
                Computername = $Env:Computername
                Date         = "-"
                ActionID     = "-"
                Action       = "-"
                Message      = $Msg
            }
        }
        Else {
            
            # Filter by date and/or action type
            If ($ArgList.ActionTypeFilter) {
                If (-not ($Output = $Output | Where-Object {$_.Action -match "^$($ArgList.ActionTypeFilter)"})) {
                    $Msg = "No results found with action filter matching '^$($ArgList.ActionTypeFilter)')"
                    $Output = [pscustomobject]@{
                        Computername = $Env:Computername
                        Date         = "-"
                        ActionID     = "-"
                        Action       = "-"
                        Message      = $Msg
                    }
                }
            }
            If ($ArgList.Start -and $ArgList.End) {
                If (-not ($Output = $Output | Where-Object {($_.Date -gt $ArgList.Start) -and ($_.Date -lt $ArgList.End)})) {
                    $Msg = "No results found between $($ArgList.Start.ToString()) and $($ArgList.End.ToString())"
                    $Output = [pscustomobject]@{
                        Computername = $Env:Computername
                        Date         = "-"
                        ActionID     = "-"
                        Action       = "-"
                        Message      = $Msg
                    }
                }
            }
        }
        
        # All done!
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

    # Hashtable for scriptblock params
    $ArgList = @{}
    $ArgList = @{
        CustomFilePath   = $CustomFilePath
        Start            = $Start
        End              = $EndDate
        ActionTypeFilter = $ActionTypeFilter
    }

    # Splat for Invoke-Command
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName   = $Null
        Authentication = "Kerberos"
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $ArgList
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.Add("AsJob",$True)
        $Param_IC.Add("JobName",$Null)
    }

    # Splat for Write-Progress
    If ($AllFiles.IsPresent) {
        $Activity = "Get content of all discovered Tanium Client action logs"    
    }
    Else {
        $Activity = "Get content of latest Tanium Client action log"    
    }
    If ($ActionTypeFilter) {
        $Activity += " matching action '^$ActionTypeFilter'"
    }
    If ($Start -and $End) {
        $Activity += " between $($Start.ToString()) and $($End.ToString())"
    }
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
    $ConfirmMsg = $Activity
    
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

} # end Get-PKTaniuamActionLog


