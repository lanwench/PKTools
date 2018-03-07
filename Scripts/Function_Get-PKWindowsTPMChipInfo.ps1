#Requires -version 3
Function Get-PKWindowsTPMChipInfo {
<# 
.SYNOPSIS
    Gets TPM chip data for a local or remote computer, interactively or as a PSJob

.DESCRIPTION
    Gets TPM chip data for a local or remote computer, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsTPMChipInfo.ps1
    Created : 2018-01-23
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-01-23 - Created script

.PARAMETER ComputerName
    Name of computer to do cool thing on; separate multiple names with commas

.PARAMETER Credential
    Valid credentials on target computer (default is current user credentials)

.PARAMETER AsJob
    Do the cool thing as a job

.PARAMETER WaitForJob
    If running as a PSJob, wait for completion and return results

.PARAMETER JobWaitTimeout
    If WaitForJob, timeout in seconds for job output

.PARAMETER SkipConnectionTest
    Don't test WinRM connectivity before submitting command

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKWindowsTPMChipInfo

        Action: Get WMI TPM chip details

        ComputerName                : SERVER11
        IsTPMChip                   : False
        IsActivated_InitialValue    : 
        IsEnabled_InitialValue      : 
        IsOwned_InitialValue        : 
        ManufacturerId              : 
        ManufacturerVersion         : 
        ManufacturerVersionInfo     : 
        PhysicalPresenceVersionInfo : 
        SpecVersion                 : 
        Messages                    : No WMI TPM chip data found


.EXAMPLE
    PS C:\> $C | Get-PKWindowsTPMChipInfo -AsJob -JobType Local -Credential $AdminCred -Verbose -OutVariable myjobs

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        AsJob                 True                                     
        JobType               Local                                    
        Credential            System.Management.Automation.PSCredential
        Verbose               True                                                                     
        ComputerName          {PKINGSLEY-04343}                        
        SkipConnectionTest    False                                    
        SuppressConsoleOutput False                                    
        ScriptName            Get-PKWindowsTPMChipInfo                 
        ScriptVersion         1.0.0                                    

        Action: Get WMI TPM chip details as local PSJob
        VERBOSE: WORKSTATION1
        VERBOSE: DEVLAPTOP1
        VERBOSE: DEVLAPTOP2
        VERBOSE: FRONTDESK
        VERBOSE: TESTVM
       
        VERBOSE: 5 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        296    TPM_WORKSTAT... BackgroundJob   Running       True            localhost            ...                      
        298    TPM_DEVLAPTOP.. BackgroundJob   Failed        False           localhost            ...                      
        300    TPM_DEVLAPTOP.. BackgroundJob   Completed     True            localhost            ...                      
        302    TPM_FRONTDESK   BackgroundJob   Running       True            localhost            ...                      
        304    TPM_TESTVM      BackgroundJob   Running       True            localhost            ...                      
        
        PS C:\> Get-Job tpm* | Where-Object {$_.State -eq "Completed"} | Receive-Job

        ComputerName                : WORKSTATION1
        IsTPMChip                   : True
        IsActivated_InitialValue    : True
        IsEnabled_InitialValue      : True
        IsOwned_InitialValue        : True
        ManufacturerId              : 1398033696
        ManufacturerVersion         : 13.12
        ManufacturerVersionInfo     : 50
        PhysicalPresenceVersionInfo : 1.2
        SpecVersion                 : 1.2, 2, 3
        Messages                    : 
        RunspaceId                  : 87fd4b25-9a0b-46b8-8f9d-82cb808fc361

        ComputerName                : DEVLAPTOP2
        IsTPMChip                   : True
        IsActivated_InitialValue    : False
        IsEnabled_InitialValue      : False
        IsOwned_InitialValue        : False
        ManufacturerId              : 1096043852
        ManufacturerVersion         : 37.19
        ManufacturerVersionInfo     : Not Supported
        PhysicalPresenceVersionInfo : 1.2
        SpecVersion                 : 1.2, 2, 3
        Messages                    : 
        RunspaceId                  : 78b9fd94-d03a-43cb-bad7-7dcf19d055f1

        ComputerName                : FRONTDESK
        IsTPMChip                   : Error
        IsActivated_InitialValue    : Error
        IsEnabled_InitialValue      : Error
        IsOwned_InitialValue        : Error
        ManufacturerId              : Error
        ManufacturerVersion         : Error
        ManufacturerVersionInfo     : Error
        PhysicalPresenceVersionInfo : Error
        SpecVersion                 : Error
        Messages                    : Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))
        RunspaceId                  : ef84f5f3-7807-4af3-ae5c-00a42c492dd0

        ComputerName                : TESTVM
        IsTPMChip                   : False
        IsActivated_InitialValue    : 
        IsEnabled_InitialValue      : 
        IsOwned_InitialValue        : 
        ManufacturerId              : 
        ManufacturerVersion         : 
        ManufacturerVersionInfo     : 
        PhysicalPresenceVersionInfo : 
        SpecVersion                 : 
        Messages                    : No WMI TPM chip data found
        RunspaceId                  : 54910535-a9ba-4c4b-9ce7-c27cad79efea

        
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
    [String[]] $ComputerName = $Env:ComputerName,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Valid credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory = $False,
        HelpMessage = "Run as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory = $False,
        HelpMessage = "Run as local or remote PSJob"
    )]
    [ValidateSet("Local","Remote")]
    [string]$JobType = "Remote",

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

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
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

    $Results = @()
    
    # Scriptblock for InvokeCommand or Start-Job
    $SCriptBlock = {
        
        Param($Computer,$Credential)
  
        $Param_GWMI = @{}
        $Param_GWMI = @{
            Class = "Win32_TPM"
            EnableAllPrivileges = $True
            Namespace = "root\CIMV2\Security\MicrosoftTpm"
            ErrorAction = "SilentlyContinue"
        }
        If ($Using:Computer) {
            $Name = $Computer
            $Param_GWMI.Add("ComputerName",$Computer)
        }
        Else {$Name = $Env:ComputerName }
        If ($Using:Credential.Username) {
            $Param_GWMI.Add("Credential",$Credential)
        }
        
        $Select = 'ComputerName','IsTPMChip','IsActivated_InitialValue','IsEnabled_InitialValue','IsOwned_InitialValue','ManufacturerId','ManufacturerVersion','ManufacturerVersionInfo','PhysicalPresenceVersionInfo','SpecVersion','Messages'
        $Props = "ComputerName","IsActivated_InitialValue","IsEnabled_InitialValue","IsOwned_InitialValue","ManufacturerId","ManufacturerVersion","ManufacturerVersionInfo","PhysicalPresenceVersionInfo","SpecVersion"

        Try {
            

            If ($TPMStatus = Get-WMIObject @Param_GWMI ) {
                
                $Output = New-Object PSObject -Property @{
                    ComputerName                = $Name
                    IsTPMChip                   = $True
                    IsActivated_InitialValue    = $TPMStatus.IsEnabled_InitialValue
                    IsEnabled_InitialValue      = $TPMStatus.IsEnabled_InitialValue
                    IsOwned_InitialValue        = $TPMStatus.IsOwned_InitialValue
                    ManufacturerId              = $TPMStatus.ManufacturerID
                    ManufacturerVersion         = $TPMStatus.ManufacturerVersion
                    ManufacturerVersionInfo     = $TPMStatus.ManufacturerVersionInfo
                    PhysicalPresenceVersionInfo = $TPMStatus.PhysicalPresenceVersionInfo
                    SpecVersion                 = $TPMStatus.SpecVersion
                    Messages                    = $Null
                }

            }
            Else {
                $Output = New-Object PSObject -Property @{
                    ComputerName                = $Name
                    IsTPMChip                   = $False
                    IsActivated_InitialValue    = $Null
                    IsEnabled_InitialValue      = $Null
                    IsOwned_InitialValue        = $Null
                    ManufacturerId              = $Null
                    ManufacturerVersion         = $Null
                    ManufacturerVersionInfo     = $Null
                    PhysicalPresenceVersionInfo = $Null
                    SpecVersion                 = $Null
                    Messages                    = "No WMI TPM chip data found"
                }
            }
        }
        Catch {
            $Output = New-Object PSObject -Property @{
                ComputerName                = $Name
                IsTPMChip                   = "Error"
                IsActivated_InitialValue    = "Error"
                IsEnabled_InitialValue      = "Error"
                IsOwned_InitialValue        = "Error"
                ManufacturerId              = "Error"
                ManufacturerVersion         = "Error"
                ManufacturerVersionInfo     = "Error"
                PhysicalPresenceVersionInfo = "Error"
                SpecVersion                 = "Error"
                Messages                    = $_.Exception.Message
            }
        }	
        
        Write-Output ($Output | Select-Object $Select)

    } #end scriptblock

    #region Splats

    # Splat for Write-Progress
    $Activity = "Get WMI TPM chip details"
    
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
        #ComputerName   = $Null
        Authentication = "Kerberos"
        ScriptBlock    = $ScriptBlock
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }

    # Paramet
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $JobPrefix = "TPM"
        Switch ($JobType) {
            Local {
                $Activity += " as local PSJob"
                $Param_SJ = @{}
                $Param_SJ = @{
                    ScriptBlock  = $ScriptBlock
                    #Credential   = $Credential
                    ArgumentList = $Null
                    Name         = $Null
                    ErrorAction  = "Stop"
                    Verbose      = $False
                }
            }
            Remote {
                $Activity += " as remote PSJob"
                $Param_IC.AsJob = $True
                $Param_IC.JobName = $Null
            }
        }#end switch
    } #end if job

    
    
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

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.CurrentOperation = $Computer
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = "$percentComplete%"
        Write-Progress @Param_WP
        
        [switch]$Continue = $False

        # If we're testing WinRM
        If (-not $SkipConnectionTest.IsPresent) {
            
             If (-not ($Computer -match "Localhost|$Env:ComputerName|127.0.0.1|")) {
                $Msg = "Test WinRM connection"
                Write-Verbose $Msg

                If ($PSCmdlet.ShouldProcess($Computer,$Msg)) {

                    $Param_WSMan.computerName = $Computer
                    If ($Null = Test-WSMan @Param_WSMan ) {
                        $Continue = $True
                    }
                    Else {
                        $Msg = "WinRM connection failed on $Computer"
                        If ($ErrorDetails = [regex]:: match($_.Exception.Message,'(?<=\<f\:Message\>).+(?=\<\/f\:Message\>)',"singleline").value.trim()) {$Msg = "$Msg`n$ErrorDetails"}
                        $Host.UI.WriteErrorLine("ERROR: $Msg")
                    }
                }
                Else {
                    $Msg = "WinRM connection test cancelled by user on $Computer"
                    $Host.UI.WriteErrorLine("$Msg on $Computer")
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
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Param_IC.ComputerName = $Computer
                    
                    If ($AsJob.IsPresent) {    
                        Switch ($JobType) {
                            Local {
                                $Param_SJ.ArgumentList = $Computer,$Credential
                                $Param_SJ.Name = "$JobPrefix`_$Computer"
                                $Job = Start-Job @Param_SJ
                                #$Msg = "Job ID $($Job.ID): $($Job.Name)"
                                #Write-Verbose $Msg
                                $Jobs += $Job
                            }
                            Remote {
                                $Job = $Null
                                $Param_IC.JobName = "$JobPrefix`_$Computer"
                                $Job = Invoke-Command @Param_IC 
                                $Msg = "Job ID $($Job.ID): $($Job.Name)"
                                Write-Verbose $Msg
                                $Jobs += $Job
                            }
                        }
                    }
                    Else {
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

} # end Function_Get-PKWindowsTPMChipInfo
