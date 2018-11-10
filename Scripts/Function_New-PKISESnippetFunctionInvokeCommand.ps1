#requires -Version 3
Function New-PKISESnippetFunctionInvokeCommand {
<#
.SYNOPSIS
    Adds a new PS ISE snippet containing a template function (function runs Invoke-Command on one or more computers interactively or as a job)

.DESCRIPTION
    Adds a new PS ISE snippet containing a template function (function runs Invoke-Command on one or more computers interactively or as a job)
    SupportsShouldProcess
    Returns a string

.NOTES
    Name    : Function_New-PKISESnippetFunctionInvokeCommand.ps1
    Created : 2017-11-28
    Author  : Paula Kingsley
    Version : 01.03.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2017-11-28 - Created script
        v01.01.0000 - 2018-03-06 - Updated scriptblock/snippet
        v01.02.0000 - 2018-10-15 - Updated, cosmetic improvements
        v01.03.0000 - 2018-10-19 - Renamed, added autodetect for author, made consistent with otheer New-PKISESnippet* 

.PARAMETER Author
    Author name

.PARAMETER AutoDetectAuthorFullName
    Attempt to match the current username to their full name via the registry & WMI

.PARAMETER Force
    Forces creation even if snippet name exists

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionInvokeCommand -Author "Ms. Scarlet" -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                                
        ---                      -----                                
        Author                   Ms. Scarlet                          
        Verbose                  True                                 
        AutoDetectAuthorFullName False                                
        Force                    False                                
        Confirm                                                       
        ScriptName               New-PKISESnippetFunctionInvokeCommand
        ScriptVersion            1.3.0                                

        VERBOSE: Setting author name to 'Ms. Scarlet'
        Action: Create ISE Snippet 'PK Invoke-Command function'
        VERBOSE: Snippet 'PK Invoke-Command function' created successfully

            Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                                        
        ----                -------------         ------ ----                                                                                        
        -a----       2018-10-19  05:56 PM          11394 PK Invoke-Command function.snippets.ps1xml     

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionInvokeCommand -AutoDetectAuthorFullName -Verbose
    VERBOSE: PSBoundParameters: 
	
    Key                      Value                                
    ---                      -----                                
    AutoDetectAuthorFullName True                                 
    Verbose                  True                                 
    Author                                                        
    Force                    False                                
    Confirm                                                       
    ScriptName               New-PKISESnippetFunctionInvokeCommand
    ScriptVersion            1.3.0                                

    VERBOSE: Setting author to current user's full name, 'Paula Kingsley'
    Action: Create ISE Snippet 'PK Invoke-Command function'
    Snippet 'PK Invoke-Command function' already exists; specify -Force to overwrite
    VERBOSE: 

        Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


    Mode                LastWriteTime         Length Name                                                                                        
    ----                -------------         ------ ----                                                                                        
    -a----       2018-10-19  05:56 PM          11394 PK Invoke-Command function.snippets.ps1xml                                                  

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionInvokeCommand -AutoDetectAuthorFullName -Force

    Action: Create ISE Snippet 'PK Invoke-Command function'

        Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


    Mode                LastWriteTime         Length Name                                                                                        
    ----                -------------         ------ ----                                                                                        
    -a----       2018-10-19  05:57 PM          11383 PK Invoke-Command function.snippets.ps1xml  


#>
[Cmdletbinding(
    DefaultParameterSetName = "Name",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Mandatory = $True,
        ParameterSetName = "Name",
        HelpMessage = "Author name"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Author,

    [Parameter(
        ParameterSetName = "Detect",
        HelpMessage = "Attempt to detect author name (currently logged-in user's full name)"
    )]
    [switch]$AutoDetectAuthorFullName,
    
    [Parameter(
        Mandatory = $False,
        HelpMessage = "Force creation of snippet even if name already exists"
    )]
    [switch]$Force
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.03.0000"
    
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
        $Msg = "This function requires the PowerShell ISE environment"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }

    #region Snippet info

    $SnippetName = "PK Invoke-Command function"
    $Description = "Snippet to create a new function to run Invoke-Command; created using New-PKISESnippetFunctionInvokeCommand"
    
    If ($AutoDetectAuthorFullName.IsPresent) {
        
        Function GetFullName {
            $SIDLocalUsers = Get-WmiObject Win32_UserProfile -EA Stop | select-Object Localpath,SID
            $UserName = (Get-WMIObject -class Win32_ComputerSystem -Property UserName -ErrorAction Stop).UserName
            $UserOnly = $UserName.Split("\")[1]
            Foreach ($Profile in $SIDLocalUsers) {
            
                # Match profile to current user
                If ($Profile.localpath -like "*$UserOnly"){

                    # Look up path in registry/Group Policy by SID
                    $SID = $Profile.sid
                    $DN = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\$SID" -ErrorAction SilentlyContinue | Select -ExpandProperty Distinguished-Name)
                    
                    If ($DN -match "(CN=)(.*?),.*") {
                        $Matches[2]        
                    }
                } #end if matching profile path
            }
        } #end function
        
        If (-not ($AuthorName = GetFullName)) {
            $Msg = "Failed to get current user's full name; please set Author manually"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
        Else {
            $Msg = "Setting author to current user's full name, '$AuthorName'"
            Write-Verbose $Msg
        }
    }
    Else {
        $AuthorName = $Author
        $Msg = "Setting author name to '$AuthorName'"
        Write-Verbose $Msg
    }

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
    Name    : Function_Do-Somethingcool.ps1
    Created : ##CREATEDATE##
    Author  : ##AUTHOR##
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - ##CREATEDATE## - Created script

.PARAMETER ComputerName
    One or more computer names

.PARAMETER Credential
    Valid credentials on target (default is current user credentials)

.PARAMETER AsJob
    Invoke command as a PSjob

.PARAMETER ConnectionTest
    Run WinRM or ping connectivity test prior to Invoke-Command, or no test (default is WinRM)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Do-SomethingCool -ComputerName foo -Verbose

        
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
        HelpMessage="One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

    [Parameter(
        HelpMessage="Valid credentials on target"
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
        HelpMessage = "Prefix for job name (default is 'Job')"
    )]
    [String] $JobPrefix = "Job",

    [Parameter(
        HelpMessage="Test to run prior to Invoke-Command - WinRM (default, using Kerberos), ping, or none)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

    [Parameter(
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
    
    #region Scriptblock for Invoke-Command

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

    #endregion Scriptblock for Invoke-Command

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
    $Activity = "Invoke scriptblock"
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
        $Param_IC.Add("AsJob",$True)
        $Param_IC.Add("JobName",$Null)
    }
    
    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
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
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
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
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            Write-Verbose $Msg
            $Jobs | Get-Job
            
        }
        Else {
            $Msg = "No jobs created"
            Write-Warning $Msg
        }
    } #end if AsJob

    Else {
        If ($Results.Count -eq 0) {
            $Msg = "No output returned"
            Write-Warning $Msg
        }
        Else {
            $Msg = "$($Results.Count) result(s) found"
            Write-Verbose $Msg
            Write-Output ($Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID)
        }
    }
}

} # end Do-SomethingCool


'@
    
    $Body = $Body.Replace("##AUTHOR##",$AuthorName)
    $Body = $Body.Replace("##CREATEDATE##",(Get-Date -f yyyy-MM-dd))

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

    # Filename
    $SnippetFile = $SnippetName +".snippets.ps1xml"

    If (($Test = Get-ISESnippet -ErrorAction SilentlyContinue | 
        Where-Object {$_.Name -eq $SnippetFile}) -and (-not $Force.IsPresent)) {
            $Msg = "Snippet '$SnippetName' already exists; specify -Force to overwrite"
            $Host.UI.WriteErrorLine("$Msg")
            Write-Verbose ($Test | Out-String)
    }    
    
    Else {
        $ConfirmMsg = "`n`n`t$Activity`n`n"
        If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$ConfirmMsg)) {
        
            Try {
                
                # Create snippet
                $Create = New-IseSnippet @Param_Snip

                If ($IsSuccess = Get-ISESnippet  -ErrorAction SilentlyContinue | 
                    Where-Object {($_.Name -eq $SnippetFile) -and ($_.CreationTime -lt (get-date))}) {
                    $Msg = "Snippet '$SnippetName' created successfully"
                    Write-Verbose $Msg
                    Write-Output $IsSuccess
                }
                Else {
                    $Msg = "Failed to create snippet"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg = "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine($Msg)
                }
            }
            Catch {
                $Msg = "Error creating snippet"
                If ($ErrorDetails = $_.Exception.Message) {$Msg = "; $ErrorDetails"}
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        Else {
            $Msg = "Snippet creation cancelled by user"
            $Host.UI.WriteErrorLine($Msg)
        }
    }

}
} #end New-PKISESnippetFunctionInvokeCommand