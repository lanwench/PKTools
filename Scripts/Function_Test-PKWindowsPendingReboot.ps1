
#requires -Version 3
function Test-PKWindowsPendingReboot {
<#
.SYNOPSIS
    Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer
    
.DESCRIPTION
    Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer
    Optionally performs a WinRM connection test before attempting job invocation
    Accepts pipeline input
    Output is a PSObject or boolean
    Returns a PSJob
    

.NOTES
    Name    : Function_Test-PKWindowsPendingReboot.ps1
    Version : 01.00.0000
    Author  : Paula Kingsley
    Created : 2016-09-16
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2017-09-12 - Created script based on Monimoy Sanyal's original (see gallery link)

.LINK
    https://gist.github.com/altrive/5329377

.LINK
    http://gallery.technet.microsoft.com/scriptcenter/Get-PendingReboot-Query-bdb79542

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
    [switch]$BooleanOutput
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
            ComputerName                         = $Env:ComputerName
            IsRebootPending                      = $InitialValue
            ComponentBasedServicingRebootPending = $InitialValue
            WindowsUpdateRebootPending           = $InitialValue
            PendingFileRenameOperation           = $InitialValue
            #SCCM                                 = $InitialValue
            Messages                             = $InitialValue
        }            
        
        If ($Bool.IsPresent) {$Select = "IsRebootPending"}
        Else {$Select = "ComputerName","IsRebootPending","ComponentBasedServicingRebootPending","PendingFileRenameOperation","PendingFileRenameOperation","Messages"}

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
                $Output.WindowsUpdatePending = $True
            }
            Else {
                $Output.WindowsUpdatePending = $False
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

         # Return the output object
         Return ($Output | Select $Select)    
        
    } #end scriptblock for remote command

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

    # Output arrays
    $Jobs = @()

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

        #If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
        If ($Computer) {

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
        $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
        Write-Verbose $Msg
        $Jobs | Get-Job
    }
    Else {
        $Msg = "No jobs created"
        $Host.UI.WriteErrorLine($Msg)
    }
}
}