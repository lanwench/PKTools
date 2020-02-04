#Requires -version 3
Function Remove-GNOpsWindowsGraphiteFolder {
<# 
.SYNOPSIS
    Invokes a scriptblock to remove the GraphitePowershell folder

.DESCRIPTION
    Invokes a scriptblock to remove the GraphitePowershell folder
    Accepts pipeline input
    Optionally tests connectivity to remote computers before invoking scriptblock
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Remove-GNOpsWindowsGraphiteFolder.ps1
    Created : 2019-12-30
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-12-30 - Created script

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER Credential
    Valid credentials on target (default is passthrough)

.PARAMETER Authentication
    WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'Graphite')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests ignored on local computer)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> Remove-GNOpsWindowsGraphite -ComputerName devsystem.domain.local -Credential (Get-Credential DOMAIN\jbloggs) -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        ComputerName   {devsystem.domain.local}
        Credential     System.Management.Automation.PSCredential
        Verbose        True                                     
        Authentication Negotiate                                
        AsJob          False                                    
        JobPrefix      Graphite                                 
        ConnectionTest WinRM                                    
        Quiet          False                                    
        ScriptName     Remove-GNOpsWindowsGraphite              
        ScriptVersion  1.0.0                                    
        PipelineInput  False                                    

        BEGIN  : Invoke scriptblock

        [devsystem.domain.local] Test WinRM connection
        [devsystem.domain.local] Invoke command

        ComputerName   : DEVSYSTEM
        ServiceFound   : True
        ServicePath    : C:\GraphitePowershellFunctions\nssm\current\win64\nssm.exe
        ServiceStopped : True
        ServiceRemoved : True
        FolderPath     : C:\GraphitePowershellFunctions
        FolderRemoved  : True
        ErrorMessages  : 

        END    : Invoke scriptblock
    
.EXAMPLE
    PS C:\> Get-ADComputer sql-prod-2 -server dc.domain.local | Remove-GNOpsWindowsGraphite -Credential $Credential -AsJob

        BEGIN  : Invoke scriptblock (as job)

        [sql-prod-2] Test WinRM connection
        [sql-prod-2] Invoke command as PSJob

        1 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name            PSJobTypeName   State       HasMoreData     Location          Command                  
        --     ----            -------------   -----       -----------     --------          -------                  
        9      Graphite_sq-... RemoteJob       Running     True            sql-prod-2.dom... ...                      

        END    : Invoke scriptblock (as job)


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
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more computer names or objects"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [object[]]$ComputerName,

    [Parameter(
        HelpMessage = "Valid credentials on target (default is passthrough)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)"
    )]
    [ValidateSet('Kerberos','Basic','Negotiate','Default','CredSSP')]
    [string]$Authentication = "Negotiate",

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Run Invoke-Command scriptblock as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Prefix for job name (default is 'Graphite')"
    )]
    [String] $JobPrefix = "Graphite",

    [Parameter(
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests are ignored on local computer)"
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
    #If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
    #    $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName
    #}
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        
        # using this object creation method & GWMI for downlevel clients    
        $Output = New-Object PSObject -Property @{
            ComputerName   = $Env:ComputerName
            FolderFound    = $False
            FolderPath     = "Error"
            FolderRemoved  = $False
            ErrorMessages  = $Null
        }
        $Select = "ComputerName","FolderPath","FolderRemoved","ErrorMessages"

        $Path = "$Env:SystemDrive:\GraphitePowershellFunctions"
        Try {

            If ($Svc = Get-Service GraphitePowerShell -EA SilentlyContinue) {
                $Null = Stop-Service GraphitePowerShell -Force -NoWait -Confirm:$False -EA Stop
                Start-Sleep -Seconds 5
            }
            #If (-not (Get-Service GraphitePowershell -EA SilentlyContinue)) {
                
            If ($FolderPath = Get-Item -Path $Path -ErrorAction SilentlyContinue) {
                $Output.FolderFound = $True
                $Output.FolderPath = $FolderPath.FullName

                Try {
                    $Null = Remove-Item $FolderPath -Recurse -Force -Confirm:$False -ErrorAction Stop 
                    $Output.FolderRemoved = $True
                }
                Catch {
                    $Output.ErrorMessages = "Failed to remove folder; $($_.Exception.Message)"
                }
            }
            Else {
                $Output.ErrorMessages = "Folder '$Path' not found"
            }
        
        }
        Catch {
            $Output.ErrorMessages = "Failed to test/stop running service; $($_.Exception.Message)"
        }
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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
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
    $Activity = "Invoke scriptblock to remove GraphitePowershell folder"
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
        [bool]$IsLocal = $True

        If ($Computer -in ($Env:ComputerName,"localhost","127.0.0.1")) {
            $IsLocal = $True
            $Continue = $True
        }
        Else {
            Switch ($ConnectionTest) {
                Default {$Continue = $True}
                Ping {
                    $Msg = "Ping computer"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
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
                WinRM {
                    $Msg = "Test WinRM connection"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
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
            }
        } #end if not local

        If ($Continue.IsPresent) {
            
            $ConfirmMsg = "`n`n`t$Activity`n`n"
            If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
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

