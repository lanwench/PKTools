#Requires -version 3
Function Get-PKWindowsRegistryDN {
<# 
.SYNOPSIS
    Gets the DistinguishedName value of an AD-joined Windows computer from its registry, interactively or as a PSJob

.DESCRIPTION
    Gets the DistinguishedName value of an AD-joined Windows computer from its registry, interactively or as a PSJob
    No ActiveDirectory module required; used for computer self-reporting
    Unless -SkipConnectionTest is specified, tests WinRM before proceeding
    SupportsShouldProcess
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsRegistryDN.ps1
    Created : 2018-01-23
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-01-23 - Created script
        v01.01.0000 - 2019-03-26 - Minor updates

.PARAMETER ComputerName
    Name of target computer (separate multiple names with commas)

.PARAMETER Credential
    Valid credentials on target computer (default is current user credentials)

.PARAMETER AsJob
    Run as a job 

.PARAMETER SkipConnectionTest
    Don't test WinRM connectivity before submitting command (skipped anway if local computer)

.PARAMETER Quiet
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\>  Get-PKWindowsRegistryDN -Verbose | FL

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Verbose               True                                     
        ComputerName          {}                        
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        WaitForJob            False                                    
        JobWaitTimeout        10                                       
        SkipConnectionTest    False                                    
        SuppressConsoleOutput False                                    
        ScriptName            Get-PKWindowsRegistryDN                  
        ScriptVersion         1.0.0                                    

        Action: Get computer DistinguishedName value from registry
        VERBOSE: PAULA-VM


        ComputerName      : PAULA-VM
        DistinguishedName : CN=PAULA-VM,OU=HQ,OU=Workstations,OU=Computers,DC=domain,DC=local
        Messages          : 


.EXAMPLE
    PS C:\> $Computers | Get-PKWindowsRegistryDN -Credential $AdminCred -AsJob

        Action: Get computer DistinguishedName value from registry as PSJob

        Id     Name             PSJobTypeName   State         HasMoreData     Location          Command                  
        --     ----             -------------   -----         -----------     --------          -------                  
        168    DN_WORKSTATION17 RemoteJob       Running       True            WORKSTATION17     ...                      
        170    DN_TESTBOX       RemoteJob       Failed        False           TESTBOX           ...                      
        172    DN_SQLDEV        RemoteJob       Running       True            SQLDEV            ...                      
        174    DN_WEBCONSOLE-BU RemoteJob       Running       True            WEBCONSOLE-BU     ...                      
        176    DN_FRONT-DESK    RemoteJob       Running       True            FRONT-DESK        ...                      
        178    DN_WORKSTATion11 RemoteJob       Completed     True            WORKSTATION11     ...                              
        196    DN_JBBILO        RemoteJob       Completed     True            JBBILO            ...          

        [...]

        PS C:\> Get-Job dn* | Select * -ExcludeProperty PSComputerName,RunspaceID

        ComputerName   DistinguishedName                                                                                                
        ------------   -----------------                                                                                                
        WORKSTATION17  CN=WORKSTATION17,OU=Seattle,OU=Workstations,OU=Computers,OU=Production,DC=domain,DC=local
        SQLDEV         CN=SQLDEV,OU=Seattle,OU=Workstations,OU=Computers,OU=Production,DC=domain,DC=local
        WEBCONSOLE-BU  CN=WEBCONSOLE-BU,OU=Indiana,OU=Servers,OU=Computers,OU=Production,DC=domain,DC...
        FRONT-DESK     CN=FRONT-DESK,OU=DeployO365,OU=Emeryville,OU=Workstations,OU=Computers,OU=Produc...
        WORKSTATION11  CN=WORKSTATION11,OU=StLouis,OU=Workstations,OU=Computers,OU=Production,DC=local,DC=com  
        JBBILO         CN=JBBILO,OU=Contractors,OU=Emeryville,OU=Workstations,OU=Computers,OU=Production,DC=domain,DC...
        
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
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)


Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"

    # Detect pipeline input
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and ((-not $ComputerName)) # -or (-not $InputObj -eq $Env:ComputerName))
    
    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"


    # If we didn't supply anything (not setting as default as parameter display won't indicate pipeline input)
    If (-not $ComputerName) {$ComputerName = $Env:ComputerName}

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    #Output
    $Results = @()
    
    #region Scriptblock

    # Scriptblock for invoke-command
    $ScriptBlock = {
        
        $ErrorActionPreference = "Stop"
        $InitialValue = "Error"
        $Output               = New-Object PSObject -Property @{
            ComputerName      = $Env:ComputerName
            DistinguishedName = $InitialValue
            Messages          = $InitialValue
        }
        $Select = "ComputerName","DistinguishedName","Messages"
        Try {
            $DN = Get-ItemProperty -path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine' -EA Stop | Select-Object -ExpandProperty Distinguished-Name -EA Stop
            $Output.DistinguishedName = $DN.ToString()
            $Output.Messages = $Null
        }
        Catch {
            $Output.Messages = $_.Exception.Message
        }
        Write-Output ($Output | Select-Object $Select)

    } #end scriptblock

    #endregion Scriptblock

    #region Splats

    # Splat for Write-Progress
    $Activity = "Get computer DistinguishedName from registry"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as PSJob"
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
        Authentication = "Kerberos"
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_IC.AsJob = $True
        $Param_IC.JobName = $Null
        $JobPrefix = "DN"
    }
    
    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}


} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        $Computer = $Computer.Trim()
        $Msg = $Computer
        
        $Current ++ 

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.CurrentOperation = $Computer
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = "$percentComplete%"
        Write-Progress @Param_WP
        
        [switch]$Continue = $False

        # If we're testing WinRM (which we will skip if it's the local computer)
        If (-not $SkipConnectionTest.IsPresent) {
                    
            If (-not ($Computer -match "Localhost|$Env:ComputerName|127.0.0.1|")) {

                $Msg = "Test WinRM connection"
                Write-Verbose "[$Computer] $Msg"

                If ($PSCmdlet.ShouldProcess($Computer,$Msg)) {

                    $Param_WSMan.computerName = $Computer
                    If ($Null = Test-WSMan @Param_WSMan ) {
                        $Continue = $True
                    }
                    Else {
                        $Msg = "WinRM connection failed"
                        If ($ErrorDetails = [regex]:: match($_.Exception.Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()) {$Msg += "`n$ErrorDetails"}
                        $Host.UI.WriteErrorLine("ERROR: [$Computer] $Msg")
                    }
                }
                Else {
                    $Msg = "WinRM connection test cancelled by user"
                    Write-Verbose "[$Computer] $Msg"
                }
            }
            Else {
                $Continue = $True
            }
        } 
        Else {
            $Continue = $True
        }
        
        If ($Continue.IsPresent) {
            
            $Msg = "Invoke command"
            Write-Verbose "[$Computer] $Msg"

            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Msg = "Job ID $($Job.ID): $($Job.Name)"
                        Write-Verbose "[$Computer] $Msg"
                        Write-Output $Job
                        $Jobs += $Job
                    }
                    Else {
                        Invoke-Command @Param_IC
                    }
                }
                Catch {
                    $Msg = "Operation failed "
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR: [$Computer] $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                Write-Verbose "[$Computer] $Msg"
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {
        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            Write-Verbose "[$Computer] $Msg"
            $Jobs | Get-Job   
        }
        Else {
            $Msg = "No jobs created"
            Write-Warning "[$Computer] $Msg"
        }
    } #end if AsJob

    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}
}

} # end Get-PKWindowsRegistryDN
