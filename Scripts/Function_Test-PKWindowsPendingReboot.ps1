#requires -Version 3
function Test-PKWindowsPendingReboot {
<#
.SYNOPSIS
    Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer
    
.DESCRIPTION
    Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer
    Optionally performs a WinRM connection test before attempting job invocation
    Accepts pipeline input
    Scriptblock output is a PSObject or boolean
    Returns a PSJob or returns job output with optional -WaitForJob switch
    
.NOTES
    Name    : Function_Test-PKWindowsPendingReboot.ps1
    Version : 01.00.0000
    Author  : Paula Kingsley
    Created : 2017-09-12
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2017-09-12 - Created script based on altrive's original (see github link)

.LINK
    https://gist.github.com/altrive/5329377

.LINK
    http://gallery.technet.microsoft.com/scriptcenter/Get-PendingReboot-Query-bdb79542

.EXAMPLE
    PS C:\> $Arr | Test-PKWindowsPendingReboot -Verbose
    # Get the pending reboot status of computer names in the pipeline

        VERBOSE: PSBoundParameters: 
	
        Key           Value                                    
        ---           -----                                    
        Verbose       True                                     
        ComputerName  {}                        
        Credential    System.Management.Automation.PSCredential
        BooleanOutput False                                    
        ScriptName    Test-PKWindowsPendingReboot              
        ScriptVersion 1.0.0                                    


        VERBOSE: testvm-1
        VERBOSE: Job ID 10: TestReboot_testvm-1_2017-09-16 10:22:03
        VERBOSE: foo
        ERROR: Connection failure on foo
        VERBOSE: testvm-2
        VERBOSE: Job ID 12: TestReboot_testvm-2_2017-09-16 10:22:06
        VERBOSE: 2 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name            PSJobTypeName   State         HasMoreData     Location    Command                  
        --     ----            -------------   -----         -----------     --------    -------                  
        10     TestReboot_t... RemoteJob       Running       True            testvm-1    ...                      
        12     TestReboot_t... RemoteJob       Running       True            testvm-2    ...                      

        [...]

        PS C:\> Get-Job 10,12 | Wait-job | Receive-job

        ComputerName                : testvm-1
        IsPendingReboot             : True
        ComponentBasedRebootPending : False
        WindowsUpdateRebootPending  : True
        PendingFileRenameOperation  : True
        Messages                    : 
        PSComputerName              : testvm-1
        RunspaceId                  : ff5104a2-9d7b-43d8-a61e-b54e0e29e51d

        ComputerName                : testvm-2
        IsPendingReboot             : True
        ComponentBasedRebootPending : False
        WindowsUpdateRebootPending  : False
        PendingFileRenameOperation  : False
        Messages                    : 
        PSComputerName              : testvm-2
        RunspaceId                  : 4d2ac96d-d085-4fd7-a890-839ef37d6a66

.EXAMPLE
    $ $Arr | Test-PKWindowsPendingReboot -BooleanOutput -WaitForJob
    # Get the pending reboot status of computer names in the pipeline, waiting for job output

        ERROR: Connection failure on foo

        IsPendingReboot PSComputerName  RunspaceId                          
        --------------- --------------  ----------                          
                   True ops-pktest-1    aa525000-1256-4cc5-81cb-500f2232e13e
                   True ops-winmgmt-2   c146e5c7-ad05-4025-b7ea-edd9155f4082
                   True PKINGSLEY-04343 dc280b9b-5dc2-4e72-ac06-5d5318d5993b


.EXAMPLE
    PS C:\> $Arr | Test-PKWindowsPendingReboot -BooleanOutput -SkipConnectionTest -Credential $Credential -Verbose
    # Get the pending reboot status of computer names in the pipeline, skipping the WinRM connection test and providing credentials

        VERBOSE: PSBoundParameters: 
	
        Key                Value                                    
        ---                -----                                    
        BooleanOutput      True                                     
        SkipConnectionTest True                                     
        Credential         System.Management.Automation.PSCredential
        Verbose            True                                     
        ComputerName       {}                        
        ScriptName         Test-PKWindowsPendingReboot              
        ScriptVersion      1.0.0                                    

        VERBOSE: testvm-1
        VERBOSE: Job ID 14: TestReboot_testvm-1_2017-09-16 10:24:43
        VERBOSE: foo
        VERBOSE: Job ID 16: TestReboot_foo_2017-09-16 10:24:43
        VERBOSE: testvm-2
        VERBOSE: Job ID 18: TestReboot_testvm-2_2017-09-16 10:24:43
        VERBOSE: 3 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output

        Id     Name            PSJobTypeName   State         HasMoreData     Location    Command                  
        --     ----            -------------   -----         -----------     --------    -------                  
        14     TestReboot_t... RemoteJob       Running       True            testvm-1    ...                      
        16     TestReboot_f... RemoteJob       Failed        False           foo         ...                      
        18     TestReboot_t... RemoteJob       Running       True            testvm-2    ...        
        
        [...]
           

        PS C:\> Get-Job 14,16,18 | Receive-job -Keep
        # Probably should've tested the connection first!

        [foo] Connecting to remote server foo failed with the following error message : WinRM cannot process the request. 
        The following error occurred while using Kerberos authentication: Cannot find the computer foo. 
        Verify that the computer exists on the network and that the name provided is spelled correctly. 
        For more information, see the about_Remote_Troubleshooting Help topic.
        At line:1 char:1
        + Get-Job 14,16,18 | Receive-job -Keep
        + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            + CategoryInfo          : OpenError: (foo:String) [], PSRemotingTransportException
            + FullyQualifiedErrorId : NetworkPathNotFound,PSSessionStateBroken


        [...]
        # Try again

        PS C:\> Get-Job 14,16,18 | Where-Object {$_.State -eq "Completed"} | Receive-Job | Format-Table

        IsPendingReboot PSComputerName RunspaceId                          
        --------------- -------------- ----------                          
                   True testvm-1       82f51e3a-dff1-4147-b995-69ae3b0234a0
                   True testvm-2       40a452ca-4a65-4708-a942-5af6ab205ace


#>
[Cmdletbinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "Medium"
)]
Param(

    [Parameter(
        Mandatory=$False,
        Position = 0,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Name of computer (separate multiple computers with commas)"
    )]
    [Alias("FQDN","HostName","Computer","VM","DNSDomainName","Name")]
    [ValidateNotNullOrEmpty()]
    [String[]] $ComputerName = $Env:ComputerName,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Valid admin credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Return boolean output only"
    )]
    [switch]$BooleanOutput,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Don't test WinRM connection before running job"
    )]
    [Switch] $SkipConnectionTest,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Wait for / get job results"
    )]
    [Switch] $WaitForJob,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Timeout (in seconds) to wait for job output"
    )]
    [ValidateRange(1,300)]
    [int] $WaitJobTimeout = 30,

    [Parameter(
        Mandatory   = $False,
        HelpMessage ="Suppress non-verbose console output"
    )]
    [switch] $SuppressConsoleOutput

)
    
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "1.00.0000"

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
    $ProgressPreference = "Continue"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }
    
    #region Scriptblock
    # Scriptblock for invoke-command
    $Scriptblock = {
        
        Param($BooleanOutput)

        # Define variables
        [switch]$Bool = $Using:BooleanOutput
        $ErrorActionPreference = "Stop"
        
        # Output object, not an ordered hashtable for backwards compatibility
        # so select array ensures property order
        $InitialValue = "Error"
        $Output = New-Object PSObject -Property @{
            ComputerName                = $Env:ComputerName
            IsPendingReboot             = $InitialValue
            ComponentBasedRebootPending = $InitialValue
            WindowsUpdateRebootPending  = $InitialValue
            PendingFileRenameOperation  = $InitialValue
            #SCCM                        = $InitialValue
            Messages                    = $InitialValue
        }            
        
        If ($Bool.IsPresent) {$Select = "IsPendingReboot"}
        Else {$Select = "ComputerName","IsPendingReboot","ComponentBasedRebootPending","WindowsUpdateRebootPending","PendingFileRenameOperation","Messages"}

        $Pending = 0
        $ErrMsg = @()

        #Test Component Based Servicing Reboot Pending
        Try {
            If (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA SilentlyContinue) { 
                $Pending ++
                $Output.ComponentBasedRebootPending = $True
            }
            Else {
                $Output.ComponentBasedRebootPending = $False
            }
        }
        Catch {
            $ErrMsg += $_.Exception.Message
        }
    
        # Test Windows Update Reboot Pending
        Try {
            If (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA SilentlyContinue) { 
                $Pending ++
                $Output.WindowsUpdateRebootPending = $True
            }
            Else {
                $Output.WindowsUpdateRebootPending = $False
            }
        }
        Catch {
            $ErrMsg += $_.Exception.Message
        }

        # Test Pending File Rename Operation
        Try {
            If (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA SilentlyContinue) { 
                $Pending ++
                $Output.PendingFileRenameOperation = $True
            }
            Else {
                $Output.PendingFileRenameOperation = $False
            }
        }
        Catch {
            $ErrMsg += $_.Exception.Message
        }

        <#
        # SCCM (not currently in use)    
        Try { 
           $SCCMClientUtil = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities" 
           $SCCMStatus = $SCCMClientUtil.DetermineIfRebootPending()

           If (($SCCMStatus -ne $Null) -and $SCCMStatus.RebootPending){
                $Pending ++
                $Output.SCCM = $True
           }
         }
         Catch{
            $ErrMsg += $_.Exception.Message
         }
         #>

         If ($Pending.Count -gt 0) {$Output.IsPendingReboot = $True}
         Else {$Output.IsPendingReboot = $False}

         If ($ErrMsg.Count -gt 0) {$Output.Messages = $ErrMsg -join("`n")}
         Else {$Output.Messages = $Null}

         # Return the output object
         Return ($Output | Select $Select)    
        
    } #end scriptblock for remote command
    #endregion Scriptblock

    #region Splats
    # Splat for test-wsman
    $Param_WSMan = @{}
    $Param_WSMan = @{
        ComputerName   = ""
        Authentication = "Negotiate"
        Credential     = $Credential
        ErrorAction    = "SilentlyContinue"
        Verbose        = $False
    }

    # Splat for invoke-command
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName = ""
        Credential   = $Credential
        Scriptblock  = $Scriptblock
        ArgumentList = $BooleanOutput
        AsJob        = $True
        JobName      = ""
        ErrorAction  = "Stop"
        Verbose      = $False
    }

    # Splat for write-progress
    If ($SkipConnectionTest.IsPresent) {$Activity = "Submit job to test pending reboot state"}
    Else {$Activity = "Test connection and submit job to test pending reboot state"}
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        PercentComplete  = ""
        CurrentOperation = "Working"
        Status           = "Working"
    }

    #endregion Splats

    # Output arrays
    $Jobs = @()

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

    If ($WaitForJob.IsPresent) {
        $Msg = "You have selected -WaitForJob, which waits $WaitJobTimeout second(s) for job results.`nIt will *not* return information on failed or running jobs. `nYou should run 'Get-Job #' to ensure you have retrieved all job information, including failures."
        Write-Warning $Msg
    }



}
Process {    

    # for progress bar    
    $Total = $ComputerName.Count
    $Current = 0
        
    # Loop through the array
    Foreach ($Computer in $ComputerName) {
                
        $Current ++
        [switch]$Continue = $False

        $Param_WP.CurrentOperation = $Computer
        $Param_WP.PercentComplete = (($Current/$Total) * 100) 
        $Param_WSMan.ComputerName = $Computer

        Write-Verbose $Computer
        Write-Progress @Param_WP

        If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
        
            If (-not $SkipConnectionTest.IsPresent) {
            
                If ($Null = Test-WSMan @Param_WSMan ) {$Continue = $True}
                Else {
                    $Msg = "Connection failure on $Computer"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    $Continue = $False
                }

            } #end if testing WinRM

            Else {$Continue = $True}
            
            If ($Continue.IsPresent) {
                
                Try {
                    $Time = (Get-Date).ToString()
                    $Param_IC.ComputerName = $Computer
                    $Param_IC.JobName = "TestReboot_$Computer`_$Time"
                    $Job = Invoke-Command @Param_IC
                    $Jobs += $Job
                    $Msg = "Job ID $($Job.ID): $($Job.Name)"
                    Write-Verbose $Msg
                }

                Catch {
                    $Msg = "Remote job creation failed for $Computer"
                    $ErrorDetails = $_.Exception.Message
                    $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
                    Continue
                }
            } #end if continue
               
        } #end if confirmed
        Else {
            $Msg = "Job cancelled by user on $Computer"
            $Host.UI.WriteErrorLine($Msg)
        } #end if cancelled

    } #end create job for each computer
    
    $Null = Write-Progress -Activity $Activity -Completed -ErrorAction SilentlyContinue -Verbose:$False

}
End {

    If ($Jobs.Count -gt 0) {
        
        If ($WaitForJob.IsPresent) {
            $Msg = "$($Jobs.Count) job(s) created; waiting $WaitForJob second(s) for output and removing completed jobs"
            Write-Verbose $Msg
            $Activity = $Msg
            Write-Progress -Activity $Activity

            Try {
                $Results = $Jobs | Get-Job | Wait-Job -Timeout $WaitJobTimeout | Receive-Job -ErrorAction SilentlyContinue
                If ($Jobs | Get-Job | Where-Object {($_.State -ne "Completed") -or ($_.HasMoreData -eq $False)}) {
                    $Msg = "Not all jobs completed within the $WaitJobTimeout-second timeout period. Please run 'Get-Job -Id #'"
                    Write-Warning $Msg
                }
                Write-Output $Results
            }
            Catch {
                $Msg = "Please check job output manually using 'Get-Job -Id #' "
                $ErrorDetails = $_.Exception.Message
                $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
            }
        }

        Else {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            Write-Verbose $Msg
            $Jobs | Get-Job
        }
        
    }
    Else {
        $Msg = "No jobs created"
        $Host.UI.WriteErrorLine($Msg)
    }
}
} #end Test-PKWindowsPendingReboot