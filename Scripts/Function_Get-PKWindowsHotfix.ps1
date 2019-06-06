#Requires -Version 3
Function Get-PKWindowsHotfix {
<# 
.SYNOPSIS
    Invokes a scriptblock to return installed Windows hotfixes using WMI (all or by KB number), interactively or as a PSJob

.DESCRIPTION
    Invokes a scriptblock to return installed Windows hotfixes using WMI (all or by KB number), interactively or as a PSJob
    Accepts pipeline input
    Uses Get-CIMInstance & Win32_QuickFixEngineering (Get-WMIObject if downlevel version)
    Optionally tests connectivity to remote computers before invoking scriptblock (defaults to WinRM test)
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsHotfix.ps1
    Created : 2019-05-23
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-05-23 - Created script

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER HotfixID
    One or more hotfix IDs (e.g., KB123456, 123456)

.PARAMETER Authentication
    Available authentication mechanism for remote connection: Default, Basic, Negotiate, Kerberos (default is Kerberos)

.PARAMETER Credential
    Valid credentials on target (default is passthrough; explicit credentials may be required depending on authentication mechanism)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as a job

.PARAMETER JobPrefix
    Prefix for job name (default is 'Hotfix')

.PARAMETER ConnectionTest
    Remote computer connectivity test prior to running Invoke-Command: WinRM with Kerberos, ping, or none (default is WinRM)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\>  Get-PKWindowsHotfix -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        Verbose        True                                     
        ComputerName   WORKSTATION14                          
        HotfixID                                                
        Authentication Kerberos                                 
        Credential     System.Management.Automation.PSCredential
        AsJob          False                                    
        JobPrefix      Hotfix                                   
        ConnectionTest WinRM                                    
        Quiet          False                                    
        PipelineInput  False                                    
        ScriptName     Get-PKWindowsHotfix                      
        ScriptVersion  1.0.0                                    

        BEGIN  : Get installed Windows hotfixes

        [WORKSTATION14] Run Invoke-Command scriptblock

        ComputerName : WORKSTATION14
        HotfixID     : KB2693643
        InstalledOn  : 2018-07-17 12:00:00 AM
        InstalledBy  : DOMAIN\jbloggs
        Name         : 
        Description  : Update
        ErrorMessage : 

        ComputerName : WORKSTATION14
        HotfixID     : KB4471331
        InstalledOn  : 2018-12-20 12:00:00 AM
        InstalledBy  : NT AUTHORITY\SYSTEM
        Name         : 
        Description  : Security Update
        ErrorMessage : 

        ComputerName : WORKSTATION14
        HotfixID     : KB4477137
        InstalledOn  : 2018-12-20 12:00:00 AM
        InstalledBy  : NT AUTHORITY\SYSTEM
        Name         : 
        Description  : Security Update
        ErrorMessage : 

        ComputerName : WORKSTATION14
        HotfixID     : KB4493464
        InstalledOn  : 
        InstalledBy  : 
        Name         : 
        Description  : Security Update
        ErrorMessage : 

        END    : Get installed Windows hotfixes

.EXAMPLE
    PS C:\> Get-PKWindowsHotfix -ComputerName patchmstr1.domain.local -HotfixID 4499175,KB9999999 -Authentication Negotiate -AsJob -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        ComputerName   patchmstr1.domain.local       
        HotfixID       {4499175, KB9999999}                       
        Authentication Negotiate                                
        Credential     System.Management.Automation.PSCredential
        AsJob          True                                     
        Verbose        True                                     
        JobPrefix      Hotfix                                   
        ConnectionTest WinRM                                    
        Quiet          False                                    
        PipelineInput  False                                    
        ScriptName     Get-PKWindowsHotfix                      
        ScriptVersion  1.0.0                                    


        BEGIN  : Get installed Windows hotfixes by ID (as job)

        [patchmstr1.domain.local] Test WinRM connection
        [patchmstr1.domain.local] Run Invoke-Command scriptblock as PSJob

        1 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        11     Hotfix_patchm... RemoteJob      Running       True            patchmstr1.domain.lo ...                      

        END    : Get installed Windows hotfixes by ID (as job)

        <snip>

        PS C:\> Get-Job 11 | Receive-Job -Keep

        ComputerName   : PATCHMSTR1
        HotfixID       : KB4499175
        IsPresent      : True
        InstalledOn    : 2019-05-26 12:00:00 AM
        InstalledBy    : NT AUTHORITY\SYSTEM
        Name           : 
        Description    : Security Update
        ErrorMessage   : 
        PSComputerName : patchmstr1.domain.local
        RunspaceId     : 0a45819f-d48f-490b-a7de-1d645a19c045

        ComputerName   : PATCHMSTR1
        HotfixID       : KB9999999
        IsPresent      : False
        InstalledOn    : 
        InstalledBy    : 
        Name           : 
        Description    : 
        ErrorMessage   : Hotfix not found
        PSComputerName : patchmstr1.domain.local
        RunspaceId     : 0a45819f-d48f-490b-a7de-1d645a19c045

.EXAMPLE
    PS C:\> Get-PKWindowsHotfix -ComputerName testsql.domain.local -HotfixID KB4483459 -ConnectionTest None -Quiet

        ComputerName : TESTSQL
        HotfixID     : KB4483459
        IsPresent    : True
        InstalledOn  : 2019-02-20 12:00:00 AM
        InstalledBy  : NT AUTHORITY\SYSTEM
        Name         : 
        Description  : Update
        ErrorMessage : 

.EXAMPLE
    PS C:\> Get-VM pk* | Get-PKWindowsHotfix -ConnectionTest None -AsJob -JobPrefix kittens -Credential $Credential -Quiet

        Id     Name               PSJobTypeName   State         HasMoreData     Location     Command                  
        --     ----               -------------   -----         -----------     --------     -------                  
        26     kittens_pktest1... RemoteJob       Running       True            pktest1      ...                      
        28     kittens_pksql-1    RemoteJob       Running       True            pksql-1      ...                      
        30     kittens_pkjump...  RemoteJob       Running       True            pkjumpbox    ...                      


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
    [Alias("Computer","Name","HostName","FQDN","DistinguishedName")]
    $ComputerName,

    [Parameter(
        HelpMessage = "One or more hotfix IDs (e.g., KB123456, 123456)"
    )]
    [String[]] $HotfixID,

    [Parameter(
        HelpMessage = "Available authentication mechanism for remote connection: Default, Basic, Negotiate, Kerberos (default is Kerberos)"
    )]
    [ValidateSet("Default","Basic","Negotiate","Kerberos")]
    [string] $Authentication = "Kerberos",

    [Parameter(
        HelpMessage = "Valid credentials on target (default is passthrough; explicit credentials may be required depending on authentication mechanism)"
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
        HelpMessage = "Prefix for job name (default is 'Hotfix')"
    )]
    [String] $JobPrefix = "Hotfix",

    [Parameter(
        HelpMessage = "Remote computer connectivity test prior to running Invoke-Command: WinRM with Kerberos, ping, or none (default is WinRM)"
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
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    #region Scriptblock for Invoke-Command
    $Scriptblock = {

        Param($SBHotfixID)
        
        $InstalledHotfix = $Null        
        $Output = @()
        $QueryStr = 'SELECT * from Win32_QuickFixEngineering'
        If ($SBHotfixID) {
            
            # Make sure they all begin with KB
            $SBHotfixArr = @()
            Foreach ($n in $SBHotfixID) {
                $n = $n.Trim()
                If (($n -notmatch "^KB\d+$") -and ($n -match "^\d+")) {$n = "KB$n"}
                $SBHotfixArr += $n
            }
            $SBHotfixArr = ($SBHotfixArr | Select-Object -Unique).Trim()

            $Select = "ComputerName","HotfixID","IsPresent","InstalledOn","InstalledBy","Name","Description","ErrorMessage"
            [string]$Where = " WHERE HotfixID LIKE '" + ($SBHotfixArr -join "' OR HotfixID LIKE '") +"'"
            $QueryStr += " $Where"
        }
        Else {
            $Select = "ComputerName","HotfixID","InstalledOn","InstalledBy","Name","Description","ErrorMessage"
        }

        Try {
            Switch ($PSVersionTable.PSVersion.Major) {
                {1..2} {
                    $InstalledHotfix = Get-WmiObject -Query $QueryStr -ErrorAction SilentlyContinue  | 
                        Select-Object @{N="ComputerName";E={$Env:ComputerName}},
                        HotfixID,
                        @{N="IsPresent";E={$True}},
                        InstalledOn,
                        InstalledBy,
                        Name,
                        Description,
                        @{N="ErrorMessage";E={$Null}}
                }
                Default {
                    $InstalledHotfix = Get-CIMInstance -Query $QueryStr -ErrorAction SilentlyContinue | 
                        Select-Object @{N="ComputerName";E={$Env:ComputerName}},
                        HotfixID,
                        @{N="IsPresent";E={$True}},
                        InstalledOn,
                        InstalledBy,
                        Name,
                        Description,
                        @{N="ErrorMessage";E={$Null}}
                }
            }
            If ($InstalledHotfix) {
                $Output += $InstalledHotfix
                If ($SBHotfixArr) {
                    Foreach ($x in ($SBHotfixArr | Where-Object {$InstalledHotfix.HotfixID -notmatch $_})) {
                        $Output += New-Object PSObject -Property @{
                            ComputerName   = $Env:ComputerName
                            HotfixID       = $x
                            IsPresent      = $False
                            InstalledOn    = $Null
                            InstalledBy    = $Null
                            Name           = $Null
                            Description    = $Null
                            ErrorMessage   = "Hotfix not found"
                        }
                    } 
                }
            }
            Else {
                $ErrorMessage = "No hotfixes found"
                If ($SBHotfixID) {$ErrorMessage = "No matching hotfixes found"}
                $Output += New-Object PSObject -Property @{
                    ComputerName   = $Env:ComputerName
                    HotfixID       = &{If ($SBHotfixArr) {$SBHotfixArr}}
                    IsPresent      = $False
                    InstalledOn    = $Null
                    InstalledBy    = $Null
                    Name           = $Null
                    Description    = $Null
                    ErrorMessage   = $ErrorMessage
                } 
            }
        }
        Catch {
            $Output += New-Object PSObject -Property @{
                ComputerName   = $Env:ComputerName
                HotfixID       = "ERROR"
                IsPresent      = $False
                InstalledOn    = "ERROR"
                InstalledBy    = "ERROR"
                Name           = "ERROR"
                Description    = "ERROR"
                ErrorMessage   = $_.Exception.Message
            } 
        }
        Write-Output $Output | Select-Object $Select
    }

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
            ErrorAction    = "Stop"
            Verbose        = $False
        }
        $Output = [pscustomobject]@{
            Success = $False
            Message = $Null
        }
        Try {
            If (Test-WSMan @Param_WSMAN) {$Output.Success = $True}
            Else {
                $Output.Message = [regex]:: match($_.Exception.Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()
            }
        }
        Catch {
            $Output.Message = [regex]:: match($_.Exception.Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()
        }
        Write-Output $Output
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
    $Activity = "Get installed Windows hotfixes"
    If ($CurrentParams.HotfixID) {$Activity += " by ID"}
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
        Authentication = $Authentication
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($HotfixID) {
        $Param_IC.Add("ArgumentList",(,$HotfixID))
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
        
        # Collect the garbage
        [system.gc]::Collect()

        # Cleanup, normalize the name if we can
        If ($Computer -is [string]) {
            $Computer = $Computer.Trim()
        }
        Elseif ($Computer -is [Microsoft.ActiveDirectory.Management.ADComputer]) {
            If ($Computer.DNSHostName) {
                $Computer = $Computer.DNSHostName
            }
            Else {$Computer = $Computer.Name}
        }
        
        # For Write-Progress
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

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        Try {
                            $TestWinRM = Test-WinRM -Computer $Computer
                            If ($TestWinRM.Success -eq $True) {$Continue = $True}
                            Else {
                                $Msg = "WinRM failure ($($TestWinRM.Message))"
                                "[$Computer] $Msg" | Write-MessageError}
                        }
                        Catch {
                            $Msg = "WinRM failure ($($TestWinRM.Message))"
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
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Msg = "Run Invoke-Command scriptblock"
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
                        Invoke-Command @Param_IC | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID
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
            $Jobs | Get-Job -Verbose:$False
            
        }
        Else {
            $Msg = "No jobs created"
            $Msg | Write-MessageError
        }
    } #end if AsJob


    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title


}

} # end Get-PKWindowsHotfix

