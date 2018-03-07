#requires -Version 3
Function New-PKISEInvokeCommandSnippet {
<#
.SYNOPSIS
    Adds a new PS ISE snippet containing a template function (function runs Invoke-Command on one or more remote computers interactively or as a job)

.DESCRIPTION
    Adds a new PS ISE snippet containing a template function (function runs Invoke-Command on one or more remote computers interactively or as a job)
    SupportsShouldProcess
    Returns a string

.NOTES
    Name    : Function_New-PKISEInvokeCommandSnippet.ps1
    Created : 2017-11-28
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2017-11-28 - Created script
        v01.01.0000 = 2018-03-06 - Updated scriptblock/snippet

.PARAMETER Force
    Forces creation even if snippet name exists

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
.EXAMPLE
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
    [Parameter(
        Mandatory = $False,
        HelpMessage = "Force creation of snippet even if name already exists"
    )]
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
    
    If (-not $PSISE) {
        $Msg = "This function requires the PowerShell ISE environtment"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }

    #region Snippet info

    $Author = "Paula Kingsley"
    $SnippetName = "Function to run Invoke-Command"
    $Description = "New-InvokeCommandFunction"

    #endregion Snippet info

    #region Here-string for function content    

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
    Created : 2018-03-05
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-03-05 - Created script

.PARAMETER ComputerName
    Name of computer to do cool thing on; separate multiple names with commas

.PARAMETER Credential
    Valid credentials on target computer (default is current user credentials)

.PARAMETER AsJob
    Do the cool thing as a job

.PARAMETER ConnectionTest
    Run WinRM or ping test prior to invoke-command, or no test (default is WinRM)

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
        Mandatory = $False,
        HelpMessage = "Run Invoke-Command scriptblock as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Test to run prior to invoke-command - WinRM (default, using Kerberos), ping, or none)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

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
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)
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
    
    # Output
    [array]$Results = @()
    
    #region Scriptblock for invoke-command

    $ScriptBlock = {
        
        Param($Arguments)

        $ErrorActionPreference = "Stop"
        $InitialValue = "Error"
        $Output = New-Object PSObject -Property @{
            ComputerName = $Env:ComputerName
            Messages     = $InitialValue
        }
        $Select = "ComputerName","Messages"
        
        Try {
            # Impressive code here
        }
        Catch {
            $Output.Messages = $_.Exception.Message
            Write-Output ($Output | Select $Select)
        }
        
        Write-Output ($Output | Select-Object $Select)

    } #end scriptblock

    #endregion Scriptblock for invoke-command

    #region Functions

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

    Function Test-Ping{
        Param($Computer)
        $Task = (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
        If ($Task.Result.Status -eq "Success") {$True}
        Else {$False}
    }

    #endregion Functions

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Write-Progress
    $Activity = "Do a cool thing"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as remote PSJob"
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
    $Msg = "[SCRIPT BEGIN] $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}


} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        $Computer = $Computer.Trim()
        $Current ++ 

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = $Computer
        
        [switch]$Continue = $False

        Switch ($ConnectionTest) {
            Default {$Continue = $True}
            Ping {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Ping computer"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-Ping -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "Ping failure"
                            $Host.UI.WriteErrorLine("[$Computer] $Msg")
                        }
                    }
                    Else {
                        $Msg = "Ping connection test cancelled by user"
                        $Host.UI.WriteErrorLine("[$Computer] $Msg")
                    }
                }
                Else {$Continue = $True}
            }
            WinRM {
                If ($Computer -ne $env:COMPUTERNAME) {
                    $Msg = "Test WinRM connection"
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n$Msg`n")) {
                        If ($Null = Test-WinRM -Computer $Computer) {$Continue = $True}
                        Else {
                            $Msg = "WinRM failure"
                            $Host.UI.WriteErrorLine("[$Computer] $Msg")
                        }
                    }
                    Else {
                        $Msg = "WinRM connection test cancelled by user"
                        $Host.UI.WriteErrorLine("[$Computer] $Msg")
                    }
                }
                Else {$Continue = $True}
            }        
        }

        If ($Continue.IsPresent) {
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
                Try {
                    $Msg = "Invoke command"
                    If ($AsJob.IsPresent) {$Msg += " as PSJob"}
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    $Param_IC.ComputerName = $Computer
                    If ($AsJob.IsPresent) {
                        $Job = $Null
                        $Param_IC.JobName = "$JobPrefix`_$Computer"
                        $Job = Invoke-Command @Param_IC 
                        $Jobs += $Job
                    }
                    Else {
                        $Results += Invoke-Command @Param_IC
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
                    $Host.UI.WriteErrorLine("[$Computer] $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                $Host.UI.WriteErrorLine("[$Computer] $Msg")
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {

        If ($Jobs.Count -gt 0) {
            $Msg = "[SCRIPT END] $($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            $FGColor = "Green"
            If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg")}
            Else {Write-Verbose $Msg}
            $Jobs | Get-Job
            
        }
        Else {
            $Msg = "[SCRIPT END]  No jobs created"
            $Host.UI.WriteErrorLine("`n$Msg")
        }
    } #end if AsJob

    Else {
        If ($Results.Count -eq 0) {
            $Msg = "[SCRIPT END] No results found"
            $Host.UI.WriteErrorLine("`n$Msg")
        }
        Else {
            $Msg = "[SCRIPT END] $($Results.Count) result(s) found"
            $FGColor = "Green"
            If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"`n$Msg")}
            Else {Write-Verbose $Msg}
            #Return $Results
            Write-Output ($Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID)
        }
    }

}

} # end Do-SomethingCool


'@

    #endregion Here-string for function content    

    #region Splat

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

    #endregion Splat

    # What are we doing
    $Activity = "Create ISE Snippet '$SnippetName'"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
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
        If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,"`n`n`t$Activity`n`n")) {
        
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