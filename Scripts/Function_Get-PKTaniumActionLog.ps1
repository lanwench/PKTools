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
    Version : 02.02.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-06-19 - Created script
        v02.00.0000 - 2019-08-13 - Added parameters for dates, action type, handling for multiple action files, local computer
        v02.01.0000 - 2019-11-14 - Fixed issue where -AllDates still asked for an end date
        v02.02.0000 - 2020-01-23 - Fixed filename filter based on changes to path/extension

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER CustomFilePath
    Absolute path to Tanium client action logfile (default is to look for action-history text file in c:\Program Files (x86)\Tanium\Tanium Client\Logs)

.PARAMETER AllFiles
    Include all discovered action-history log files (default is the most recently written to)

.PARAMETER AllDates
    Return entries from all dates (ignores start/end dates)

.PARAMETER ActionTypeFilter
    Return only actions where the filter string matches an action type (default is no filter)

.PARAMETER StartDate
    Start date (default is past 24 hours)

.PARAMETER EndDate
    End date (default is now)

.PARAMETER DateSort
    Sort return output by Ascending or Descending date (default is Descending)

.PARAMETER Credential
    Valid credentials on target (default is passthrough; credential is ignored on local computer)

.PARAMETER Authentication
    WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; authentication is ignored on local computer)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'TaniumActivity')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM with Kerberos, ping, or none (default is WinRM)

.PARAMETER NoProgress
    Don't display progress bar (can improve performance with large files)

.PARAMETER Quiet
    Hide all non-verbose console output (can improve performance with large files)

.EXAMPLE
    PS C:\> Get-PKTaniumActionLog -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                    
        ---              -----                                    
        StartDate        2019-09-04 10:15:39 AM                              
        EndDate          2019-09-05 10:15:39 AM                               
        Verbose          True                                     
        ComputerName     LAPTOP                          
        CustomFilePath                                            
        ActionTypeFilter                                          
        AllFiles         False                                    
        AllDates         False                                    
        DateSort         Descending                               
        Credential       System.Management.Automation.PSCredential
        Authentication   Negotiate                                
        ConnectionTest   WinRM                                    
        AsJob            False                                    
        JobPrefix        TaniumActivity                           
        NoProgress       False                                    
        Quiet            False                                    
        ParameterSetName Interactive                              
        PipelineInput    False                                    
        ScriptName       Get-PKTaniumActionLog                    
        ScriptVersion    2.0.0                                 


        BEGIN  : Get content of latest Tanium Client action log between 2019-09-04 10:15:39 AM and 2019-09-05 10:15:39 AM


        [LAPTOP] Invoke command

        Computername : LAPTOP
        Date         : 2019-09-04 11:59:12 PM
        ActionID     : 1578559
        Action       : Windows
        Message      : executing: cmd /c cscript install-config.vbs /Level:0

        Computername : LAPTOP
        Date         : 2019-09-04 11:56:05 PM
        ActionID     : 1578632
        Action       : Clean Stale Tanium Client Data
        Message      : completed in 2.884 seconds with exit code 0

        Computername : LAPTOP
        Date         : 2019-09-04 11:56:02 PM
        ActionID     : 1578632
        Action       : Clean Stale Tanium Client Data
        Message      : executing: cmd /c cscript //T:1200 clean-stale-tanium-client-data.vbs

        Computername : LAPTOP
        Date         : 2019-09-04 11:44:39 PM
        ActionID     : 1578586
        Action       : Deploy - Deployment 86
        Message      : completed in 1.238 seconds with exit code 0

        Computername : LAPTOP
        Date         : 2019-09-04 11:44:38 PM
        ActionID     : 1578586
        Action       : Deploy - Deployment 86
        Message      : executing: cmd /c xcopy /y deployment_86.json ..\..\Tools\Deploy\deployments\configurations\

        Computername : LAPTOP
        Date         : 2019-09-04 11:43:01 PM
        ActionID     : 1578578
        Action       : Patch - Blacklist 4
        Message      : executing: cmd /c xcopy /y blacklist-4.xml ..\..\Patch\blacklists\configurations\

        Computername : LAPTOP
        Date         : 2019-09-04 11:40:26 PM
        ActionID     : 1578557
        Action       : Patch - Scan Configuration Priorities - Windows
        Message      : completed in 0.170 seconds with exit code 0

        END    : Get content of latest Tanium Client action log between 2019-09-04 10:15:39 AM and 2019-09-05 10:15:39 AM


.EXAMPLE
    PS C:\> Get-PKTaniumActionLog -ComputerName sqldev2.domain.local -ActionTypeFilter Detect -Credential $Credential -Authentication Kerberos -ConnectionTest None -AsJob -JobPrefix test  -Verbose 

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                    
        ---              -----                                    
        ComputerName     {sqldev2.domain.local}
        Credential       System.Management.Automation.PSCredential
        Authentication   Kerberos                                 
        AsJob            True                                     
        JobPrefix        test                                     
        ActionTypeFilter Detect                                   
        Verbose          True                                     
        CustomFilePath                                            
        AllFiles         False                                    
        AllDates         False                                    
        StartDate        2019-09-08 10:01:49 AM                   
        EndDate          2019-09-09 10:01:49 AM                   
        DateSort         Descending                               
        ConnectionTest   None
        NoProgress       False                                    
        Quiet            False                                    
        ParameterSetName Job                                      
        PipelineInput    False                                    
        ScriptName       Get-PKTaniumActionLog                    
        ScriptVersion    2.0.0                                    

        BEGIN  : Get content of latest Tanium Client action log matching action 'Detect' between 2019-09-08 10:01:49 AM and 2019-09-09 10:01:49 AM (as job)

        [sqldev2.domain.local] Invoke command as PSJob

        1 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output


        Id     Name            PSJobTypeName   State    HasMoreData   Location       Command                  
        --     ----            -------------   -----    -----------   --------       -------                  
        61     test_sqldev2... RemoteJob       Running  True          sqldev2.dom... ...                     

        END    : Get content of latest Tanium Client action log matching action 'Detect' between 2019-09-08 10:01:49 AM and 2019-09-09 10:01:49 AM (as job)


        PS C:\> Get-Job 61 | Receive-Job

        Computername : SQLDEV2
        Date         : 2019-09-09 09:39:43 PM
        ActionID     : 1577895
        Action       : Detect Intel for Windows Revision 753 Sync
        Message      : completed in 0.347 seconds with exit code 0

        Computername : SQLDEV2
        Date         : 2019-09-09 09:39:42 PM
        ActionID     : 1577895
        Action       : Detect Intel for Windows Revision 753 Sync
        Message      : executing: cmd /c cscript /nologo run-add-intel-package.vbs 2>&1

        Computername : SQLDEV2
        Date         : 2019-09-08 11:12:13 PM
        ActionID     : 1578327
        Action       : Detect Intel for Windows Revision 753 Sync
        Message      : completed in 0.287 seconds with exit code 0

        Computername : SQLDEV2
        Date         : 2019-09-08 11:12:13 PM
        ActionID     : 1578327
        Action       : Detect Intel for Windows Revision 753 Sync
        Message      : executing: cmd /c cscript /nologo run-add-intel-package.vbs 2>&1

        Computername : SQLDEV2
        Date         : 2019-09-08 10:24:55 PM
        ActionID     : 1578163
        Action       : Detect Group Config 816 revision 2 - ALL_WINDOWS - Windows
        Message      : completed in 1.995 seconds with exit code 0

        Computername : SQLDEV2
        Date         : 2019-09-08 10:24:53 PM
        ActionID     : 1578163
        Action       : Detect Group Config 816 revision 2 - ALL_WINDOWS - Windows
        Message      : executing: cmd /c cscript /nologo deploy-detect-tools.vbs


.EXAMPLE
    PS C:\> Get-PKTaniumActionLog -StartDate 2019-09-05 -AsJob

        BEGIN  : Get content of latest Tanium Client action log between 2019-09-05 12:00:00 AM and 2019-09-09 09:53:41 AM (as job)

        [LAPTOP] Invoke command as PSJob

        1 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command               
        --     ----            -------------   -----         -----------     --------             -------               
        59     TaniumActivi... BackgroundJob   Running       True            localhost            ...                   

        END    : Get content of latest Tanium Client action log between 2019-09-05 12:00:00 AM and 2019-09-09 09:53:41 AM (as job)


        PS C:\> Get-Job 59 | Receive-Job | Select-Object -Property * -ExcludeProperty RunspaceID,PSComputerName | Format-Table -AutoSize

        Computername  Date                   ActionID Action                             Message                                                                                                  
        ------------  ----                   -------- ------                             -------                                                                                                  
        LAPTOP        2019-09-06 05:25:00 PM 1591260  Deploy - Profile 2                 completed in 2.135 seconds with exit code 0...                                                           
        LAPTOP        2019-09-06 05:24:57 PM 1591260  Deploy - Profile 2                 executing: cmd /c xcopy /y profile_2.json ..\..\Tools\Deploy\profiles\...                                
        LAPTOP        2019-09-06 05:24:45 PM 1591258  Patch - Scan Configuration 2       completed in 0.580 seconds with exit code 0...                                                           
        LAPTOP        2019-09-06 05:24:45 PM 1591258  Patch - Scan Configuration 2       executing: cmd /c xcopy /y scan-configuration-2.xml ..\..\Patch\scans\configurations\...                 
        LAPTOP        2019-09-06 05:22:36 PM 1591249  Deploy - Deployment 119            completed in 1.297 seconds with exit code 0...                                                           
        LAPTOP        2019-09-06 05:22:35 PM 1591249  Deploy - Deployment 119            executing: cmd /c xcopy /y deployment_119.json ..\..\Tools\Deploy\deployments\configurations\...         
        LAPTOP        2019-09-06 05:22:35 PM 1591244  Patch - Maintenance Window 38      completed in 0.270 seconds with exit code 0...                                                           
        LAPTOP        2019-09-06 05:22:35 PM 1591244  Patch - Maintenance Window 38      executing: cmd /c xcopy /y maintenance-window-38.xml ..\..\Patch\maintenance-windows\configurations\...  
        LAPTOP        2019-09-06 05:22:12 PM 1591231  Patch - Run Patch Manifest Cleanup completed in 0.444 seconds with exit code 0...                                                           
        LAPTOP        2019-09-06 05:22:12 PM 1591231  Patch - Run Patch Manifest Cleanup executing: cmd.exe /c cscript.exe //E:VBScript //T:60 ..\..\Tools\PatchExt\cleanup-patch-manifests.vbs...
    
#> 

[CmdletBinding(
    DefaultParameterSetName = "Interactive",
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
        HelpMessage = "Return only actions where the filter string matches an action type (default is no filter)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Asset","Clean","Collect","Comply","Delete","Deploy","Detect","End-User","Execute","Patch","Run","UserGroupTagging","Trace","Windows")]
    [string] $ActionTypeFilter,
    
    [Parameter(
        HelpMessage = "Include all discovered action-history log files (default is the most recently written)"
    )]
    [Switch] $AllFiles,

    [Parameter(
        HelpMessage = "Return entries from all dates (ignores start/end dates)"
    )]
    [switch]$AllDates,

    [Parameter(
        HelpMessage = "Start date (default is past 30 days; ignored if -AllDates is present)"
    )]
    #[ValidateNotNullOrEmpty()]
    #$StartDate = ([System.DateTime]::Now).AddHours(-24),
    $StartDate = ([System.DateTime]::Now).AddDays(-30),

    [Parameter(
        HelpMessage = "End date (default is now)"
    )]
    [ValidateNotNullOrEmpty()]
    $EndDate = [System.DateTime]::Now,

    [Parameter(
        HelpMessage = "Sort return output by Ascending or Descending date (default is Descending)"
    )]
    [ValidateSet("Ascending","Descending")]
    $DateSort = "Descending",

    [Parameter(
        HelpMessage = "Valid credentials on target (default is passthrough; credential is ignored on local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; authentication is ignored on local computer)"
    )]
    [ValidateSet('Kerberos','Basic','Negotiate','Default','CredSSP')]
    [string]$Authentication = "Negotiate",

    [Parameter(
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

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
        HelpMessage = "Don't display progress bar (can improve performance with large files)"
    )]
    [Switch] $NoProgress,

    [Parameter(
        HelpMessage = "Hide all non-verbose console output (can improve performance with large files)"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "02.02.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    If ($AllDates.IsPresent) {
        $CurrentParams.StartDate = $Null
    }
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName
    }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Preferences 
    
    $ErrorActionPreference = "Stop"
    Switch ($NoProgress) {
        $True  {$ProgressPreference = "SilentlyContinue"}
        $False {$ProgressPreference = "Continue"}
    }

    #endregion Preferences
    
    #region Parameter things

    If ($AllDates.IsPresent -and $StartDate) {
        $Msg = "You have selected -AllDates; StartDate and EndDate values will be ignored"
        Write-Warning $Msg
        $StartDate = ""
    }
    Else {
        If ($StartDate -and $EndDate) {
            If (-not ($StartDate -as [datetime])) {
                $Msg = "'$StartDate' is not a valid date/time; please re-enter StartDate"
                $Host.UI.WriteErrorLine($Msg)
                Break
            }
            If (-not ($EndDate -as [datetime])) {
                $Msg = "'$EndDate' is not a valid date/time; please re-enter EndDate"
                $Host.UI.WriteErrorLine($Msg)
                Break
            }
            Else {
                $Start = ($StartDate -as [datetime])
                $End = ($EndDate -as [datetime])
                If ($Start -gt (Get-Date)) {
                    $Msg = "Time travel is not permitted; please enter a start date earlier than $((Get-Date).ToString())"
                    $Host.UI.WriteErrorLine($Msg)
                    Break
                }
                If (-not ($End -ge $Start)) {
                    $Msg = "Time travel is not permitted; please enter an end date later than or equal to $(($StartDate -as [datetime]).ToString())"
                    $Host.UI.WriteErrorLine($Msg)
                    Break
                }
            }
        }
        Else {$StartDate = $EndDate = ""}
    }
    
    #endregion Parameter things

    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        
        Param($ArgList)

        If ($ArgList.CustomFilePath) {
            If (Test-Path $ArgList.CustomFilePath -PathType Container -EA SilentlyContinue) {
                $Path = $ArgList.CustomFilePath
                Try {
                    $ActionFile = Get-ChildItem $CustomFilePath -Recurse -Filter action-history*.txt -EA SilentlyContinue -Force | Get-Item -Force
                }
                Catch {
                    $Msg = "No Tanium client activity logfile found in path '$Path'"
                }
            }
            Else {
                $Msg = "Invalid path '$($ArgList.CustomFilePath)"
            }
        }
        Else {
            $TaniumCommandPath = "$($Env:SystemDrive)\Program Files*\Tanium\Tanium Client\taniumclient.exe"
            $Path = ($TaniumCommandPath | Select-Object $_.Path | Split-Path -Parent)

            Try {
                $Parent = Get-Command -Name $TaniumCommandPath -EA Stop | Select-Object $_.Path | Split-Path -Parent -EA Stop
                #$ActionFile = Get-ChildItem $Parent -Recurse -Filter action-history*.txt -Force -EA Stop | Get-Item -Force -EA Stop
                $FileRegex = "\.log|\.txt"
                $ActionFile = Get-ChildItem $Parent -Recurse -Filter "action*" -Force -EA Stop | Where-Object -FilterScript {$_.Name -match $FileRegex}
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
                            New-Object PSObject -Property @{
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
        Else {
            $Msg = "No Tanium action-history file(s) found in path '$Path'"
        }
            
        If (-not $Output) {
            $Output = New-Object PSObject -Property @{
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
                    $Output = New-Object PSObject -Property @{
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
                    $Output = New-Object PSObject -Property @{
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
        If ($ArgList.DateSort -eq "Descending") {
            Write-Output $Output | Select-Object ComputerName,Date,ActionID,Action,Message | Sort-Object Date -Descending
        }
        Else {
            Write-Output $Output | Select-Object ComputerName,Date,ActionID,Action,Message | Sort-Object Date
        }
        

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

    # Hashtable for scriptblock params
    $ArgList = @{}
    $ArgList = @{
        CustomFilePath   = $CustomFilePath
        Start            = $Start
        End              = $End
        ActionTypeFilter = $ActionTypeFilter
        DateSort         = $DateSort
    }

    # Splats for Invoke-Command (remote)
    $Param_IC_Remote = @{}
    $Param_IC_Remote = @{
        ComputerName   = $Null
        Authentication = $Authentication
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $ArgList
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_Job = @{}
        $Param_Job = @{
            AsJob   = $True
            JobName = $Null
        }
    }

    # Splat for Invoke-Command (local)
    $Param_IC_Local = @{}
    $Param_IC_Local = @{
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $ArgList
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    
    
    # Splat for Start-Job (local computer)
    $Param_SJ = @{}
    $Param_SJ = @{
        Authentication = "Default"
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $ArgList
        Name           = $Null
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($Credential.Username) {
        $Param_SJ.Add("Credential",$Credential)
    } 

    # Splat for Write-Progress
    If ($AllFiles.IsPresent) {
        $Activity = "Get content of all discovered Tanium Client action logs"    
    }
    Else {
        $Activity = "Get content of latest Tanium Client action log"    
    }
    If ($ActionTypeFilter) {
        $Activity += " matching action '$ActionTypeFilter'"
    }
    If (-not $AllDates.IsPresent -and ($Start -and $End)) {
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
                    If ($Env:ComputerName -eq $Computer) {
                        If ($AsJob.IsPresent) {
                            $Job = $Null
                            $Param_SJ.Name = "$JobPrefix`_$Computer"
                            $Job = Start-Job @Param_SJ
                            $Jobs += $Job
                        }
                        Else {
                            Invoke-Command @Param_IC_Local | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceId
                        }
                    }
                    Else {   
                        $Param_IC_Remote.ComputerName = $Computer
                        If ($AsJob.IsPresent) {
                            $Job = $Null
                            $Param_Job.JobName = "$JobPrefix`_$Computer"
                            $Job = Invoke-Command @Param_IC @Param_Job
                            $Jobs += $Job
                        }
                        Else {
                            Invoke-Command @Param_IC | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceId
                        }   
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
            "$Msg" | Write-MessageInfo -FGColor Green -Title
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


