#requires -Version 3
Function New-PKISEInvokeCommandSnippet {
<#
.SYNOPSIS
    Adds a new PS ISE snippet containing a template function to run Invoke-Command on one or more remote computers (optionally as a job)
    
.DESCRIPTION
    Adds a new PS ISE snippet containing a template function to run Invoke-Command on one or more remote computers (optionally as a job)
    Returns a string
    
.NOTES
    Name    : Function_New-PKISEInvokeCommandSnippet.ps1
    Version : 01.00.0000
    Author  : Paula Kingsley
    Created : 2017-11-28
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2017-11-28 - Created script

.EXAMPLE
    PS C:\> New-PKISEInvokeCommandSnippet -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                        
        ---           -----                        
        Verbose       True                         
        Force         False                        
        ScriptName    New-PKISEInvokeCommandSnippet
        ScriptVersion 1.0.0                        

        ACTION: Create ISE Snippet 'Function to run Invoke-Command'
        VERBOSE: Snippet 'Function to run Invoke-Command' created successfully


            Directory: C:\Users\pkingsley\documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                           
        ----                -------------         ------ ----                                                                           
        -a----       2017-11-28     12:07          12156 Function to run Invoke-Command.snippets.ps1xml  

.EXAMPLE
    PS C\> New-PKISEInvokeCommandSnippet -Verbose

    VERBOSE: PSBoundParameters: 
	
        Key           Value                        
        ---           -----                        
        Verbose       True                         
        Force         False                        
        ScriptName    New-PKISEInvokeCommandSnippet
        ScriptVersion 1.0.0                        

        ACTION: Create ISE Snippet 'Function to run Invoke-Command'
        ERROR: Snippet 'Function to run Invoke-Command' already exists; specify -Force to overwrite
        VERBOSE: 

            Directory: C:\Users\pkingsley\documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                           
        ----                -------------         ------ ----                                                                           
        -a----       2017-11-28     12:10          12156 Function to run Invoke-Command.snippets.ps1xml     
.EXAMLE
    PS C:\> New-PKISEInvokeCommandSnippet -Force -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                        
        ---           -----                        
        Force         True                         
        Verbose       True                         
        ScriptName    New-PKISEInvokeCommandSnippet
        ScriptVersion 1.0.0                        

        ACTION: Create ISE Snippet 'Function to run Invoke-Command'
        VERBOSE: Snippet 'Function to run Invoke-Command' created successfully

            Directory: C:\Users\pkingsley\documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                           
        ----                -------------         ------ ----                                                                           
        -a----       2017-11-28     12:10          12156 Function to run Invoke-Command.snippets.ps1xml  

        
#>
[Cmdletbinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [switch]$Force
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
    $ProgressPreference = "Continue"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }
    
    $Author = "Paula Kingsley"
    $SnippetName = "Function to run Invoke-Command"
    $Description = "New-InvokeCommandFunction"


    If (-not $PSISE) {
        $Msg = "This function requires the PowerShell ISE environtment"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }

# Here-string for function content    
$Body = @'
#Requires -version 3
Function Do-SomethingCool {
<# 
.SYNOPSIS
    Does something cool, interactively or as a PSJob

.DESCRIPTION
    Does something cool, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_do-Somethingcool.ps1
    Created : 2017-11-28
    Author  : Paula Kingsley
    Version : 02.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2017-11-28 - Created script

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
    PS C:\> Do-SomethingCool -ComputerName foo

        
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
        Mandatory=$True,
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
        ParameterSetName = "Job",
        Mandatory=$False,
        HelpMessage="Wait for job completion"
    )]
    [Switch] $WaitForJob,

    [Parameter(
        ParameterSetName = "Job",
        Mandatory=$False,
        HelpMessage="Timeout in seconds to wait for job results (default 10)"
    )]
    [int] $JobWaitTimeout = 10,

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
    
    # Scriptblock for invoke-command
    $SCriptBlock = {
        
        Param($Arguments)

        $ErrorActionPreference = "Stop"

        $Results = @()
        
        $InitialValue = "Error"
        $OutputTemplate = @{
            ComputerName = $Env:ComputerName
            Messages     = $InitialValue
        }

        $Select = "ComputerName","Messages"
        
        
        Try {
            
            # Impressive code here
        }
        Catch {
            $OutputTemplate.Messages = $_.Exception.Message
            $Results += New-Object PSObject -property $OutputTemplate
        }
        
        Write-Output ($Results | Select-Object $Select)

    } #end scriptblock

    #region Splats

    # Splat for Write-Progress
    $Activity = "Do a cool thing"
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
        $JobPrefix = "Thing"
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

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.CurrentOperation = $Computer
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = "$percentComplete%"
        Write-Progress @Param_WP
        
        [switch]$Continue = $False

        # If we're testing WinRM
        If (-not $SkipConnectionTest.IsPresent) {
                    
            $Msg = "Test WinRM connection"
            Write-Verbose $Msg

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
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Msg = "Job ID $($Job.ID): $($Job.Name)"
                        Write-Verbose $Msg
                        $Jobs += $Job
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
        
            If ($WaitForJob.IsPresent) {
                $Msg = "$($Jobs.Count) job(s) created (waiting $JobWaitTimeout second(s) for output)"
                Write-Verbose $Msg
                $Activity = $Msg
                Write-Progress -Activity $Activity
                $FGColor = "Yellow"
                If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
                Else {Write-Verbose $Msg}

                Start-Sleep 5

                Try {
                    $Results = @()
                    If ($NotComplete = ($Jobs | Get-Job | Where-Object {(@("Failed","Running") -contains $_.State)})) {
                        $Msg = "Not all jobs completed within the $JobWaitTimeout-second timeout period. Please run 'Get-Job -Id #'`n"
                        Write-Warning $Msg
                        Write-Verbose ($Notcomplete | Get-Job @StdParams | Out-String)
                    }
                    If (($Results = $Jobs | Get-Job | Wait-Job -Timeout $JobWaitTimeout | Receive-Job -ErrorAction SilentlyContinue -Verbose:$False).Count -gt 0) {
                        Write-Output ($Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID)
                    }
                    Else {
                        $Msg = "No job output returned (try -IncludeNoMatch to return data for computers where no matching services were found)"
                        Write-Warning $Msg
                    }
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

} # end Do-SomethingCool

'@

    # Splat
    $Param_Snip = @{
        Title       = $SnippetName
        Description = $Description
        Text        = $Body
        Author      = $Author
        Verbose     = $False
        ErrorAction = "Stop"
    }
    If ($CurrentParams.ContainsKey("Force")) {
        $Param_Snip.Add("Force",$Force)
    }

    # What are we doing
    $Activity = "Create ISE Snippet '$SnippetName'"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "ACTION: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg} 
}
Process {

    If (($Test = Get-ISESnippet -ErrorAction SilentlyContinue | Where-Object {$_.Name -eq $SnippetName +".snippets.ps1xml"}) -and (-not $Force.IsPresent)) {
        $Msg = "Snippet '$SnippetName' already exists; specify -Force to overwrite"
        $Host.UI.WriteErrorLine("$Msg")
        Write-Verbose ($Test | Out-String)
    }    
    
    Else {
        If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$Activity)) {
        
            Try {
                New-IseSnippet @Param_Snip
                If ($Test = Get-ISESnippet -ErrorAction SilentlyContinue | Where-Object {$_.Name -eq $SnippetName +".snippets.ps1xml" -and ((Get-Date).ToShortDateString() -match $_.CreationTime.ToShortDateString())}) {
                    $Msg = "Snippet '$SnippetName' created successfully"
                    Write-Verbose $Msg
                    Write-Output $Test
                }
                Else {
                    $Msg = "Snippet creation failed"
                    $Host.UI.WriteErrorLine($Msg)
                }
            }
            Catch {
                $Msg = "Snippet creation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        Else {
            $Msg = "Snippet creation cancelled by user"
            $Host.UI.WriteErrorLine($Msg)
        }
    }

}
} #end New-PKISEInvokeCommandSnippet