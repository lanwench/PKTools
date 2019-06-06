#requires -Version 3
Function New-PKISESnippetFunctionInvokeCommand {
<#
.SYNOPSIS
    Adds a new PS ISE snippet containing a template function (function runs Invoke-Command on one or more computers interactively or as a job)

.DESCRIPTION
    Adds a new PS ISE snippet containing a template function (function runs Invoke-Command on one or more computers interactively or as a job)
    Author name can be manually specified or the current user's name can be detected using WMI & the local registry
    SupportsShouldProcess
    Returns a string

.NOTES
    Name    : Function_New-PKISESnippetFunctionInvokeCommand.ps1
    Created : 2017-11-28
    Author  : Paula Kingsley
    Version : 01.04.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2017-11-28 - Created script
        v01.01.0000 - 2018-03-06 - Updated scriptblock/snippet
        v01.02.0000 - 2018-10-15 - Updated, cosmetic improvements
        v01.03.0000 - 2018-10-19 - Renamed, added autodetect for author, made consistent with other New-PKISESnippet* 
        v01.04.0000 - 2019-05-23 - Updated scriptblock for snippet & changed parameter for autodetect v manual authorname,
                                   other cosmetic changes

.PARAMETER Author
    Author name

.PARAMETER AuthorNameOption
    Set author name manually, or attempt to detect current user's full name (default)

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
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    
    [Parameter(
        HelpMessage = "Set author name manually, or attempt to detect current user's full name (default)"
    )]
    [ValidateSet("Manual","AutoDetect")]
    [String]$AuthorNameOption = "AutoDetect",
    
    [Parameter(
        HelpMessage = "Author name (if not using Autodetect)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Author,

    [Parameter(
        HelpMessage = "Overwrite existing snippet"
    )]
    [switch]$Force,

    [Parameter(
        HelpMessage = "Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.04.0000"
    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    If (($AuthorNameOption -eq "Autodetect") -and ($CurrentParams.Author)) {
        $AuthorNameOption = "Manual"
        $CurrentParams.AuthorNameOption = "Manual"
    }
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

    # Function to get the current user's full name via WMI/registry
    If ($AuthorNameOption -eq "AutoDetect") {
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
    }

    #endregion Functions

    #region Prerequisites

    If (-not $PSISE) {
        $Msg = "This function requires the PowerShell ISE environment"
        "[$Env:ComputerName] $Msg" | Write-MessageError
        Break
    }

    If ($AuthorNameOption -eq "Manual" -and (-not $CurrentParams.Author)) {
        $Msg = "Author name must be provided when -AuthorNameOption is set to 'Manual'"
        "[$Env:ComputerName] $Msg" | Write-MessageError
        Break
    } 

    #region Prerequisites

    #region Snippet variables

    $SnippetName = "PK Invoke-Command function"
    $Description = "Snippet to create a new function to run Invoke-Command; created using New-PKISESnippetFunctionInvokeCommand"
    $SnippetFile = $SnippetName +".snippets.ps1xml"

    #endregion Snippet variables

    #region Here-string for function content    

$Body = @'
#Requires -version 3
Function Do-SomethingCool {
<# 
.SYNOPSIS
    Invokes a scriptblock to do something cool, interactively or as a PSJob

.DESCRIPTION
    Invokes a scriptblock to do something cool, interactively or as a PSJob
    Accepts pipeline input
    Optionally tests connectivity to remote computers before invoking scriptblock
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Do-SomethingCool.ps1
    Created : ##CREATEDATE##
    Author  : ##AUTHOR##
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - ##CREATEDATE## - Created script

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER Credential
    Valid credentials on target (default is passthrough)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'Job')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM with Kerberos, ping, or none (default is WinRM)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> 

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
    [Alias("Computer","Name","HostName","FQDN")]
    [object[]]$ComputerName,

    [Parameter(
        HelpMessage = "Valid credentials on target (default is passthrough)"
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
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM with Kerberos, ping, or none (default is WinRM)"
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
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        
        Param($Thing)
        Try {
            Get-Item $Thing
        }
        Catch {
            Throw $_.Exception.Message
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
    $Activity = "Invoke scriptblock"
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
            
            If ($PSCmdlet.ShouldProcess($Computer,$Activity)) {
                
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
                        Invoke-Command @Param_IC
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


'@
    
    #endregion Here-string for function content    

    #region Splat

    $Param_Snip = @{
        Title       = $SnippetName
        Description = $Description
        Text        = $Null
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
    $Msg = "BEGIN: $Activity"
    "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor Yellow -Title
}
Process {
    
    
    # See if it already exists

    Write-Progress -Activity $Activity -Status $Env:ComputerName -CurrentOperation 'Test for existing snippet' -PercentComplete (1/3 * 100)

    If (($Test = Get-ISESnippet -ErrorAction SilentlyContinue | 
        Where-Object {$_.Name -eq $SnippetFile}) -and (-not $Force.IsPresent)) {
            $Msg = "ERROR: Snippet '$SnippetName' already exists; specify -Force to overwrite"
            "[$Env:ComputerName] $Msg" | Write-MessageError
            ($Test | Out-String) | Write-MessageInfo -FGColor White
    }    
    
    # Create it
    Else {
        
        Write-Progress -Activity $Activity -Status $Env:ComputerName -CurrentOperation 'Author name' -PercentComplete (2/3 * 100)
        
        If ($AuthorNameOption -eq "Autodetect") {
        
            If (-not ($AuthorName = GetFullName)) {
                $Msg = "ERROR: Failed to match current username '$Env:UserName' to a full name using the Users path, WMI/registry; please set Author manually"
                "[$Env:ComputerName] $Msg" | Write-MessageError
                Break
            }
            Else {
                $Msg = "Setting author to current user's full name, '$AuthorName'"
                "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor Green
            }
        }
        Else {
            $AuthorName = $Author
            $Msg = "Setting author name to '$AuthorName'"
            "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor Green
        }

        #endregion Snippet info

        Write-Progress -Activity $Activity -Status $Env:ComputerName -CurrentOperation 'Create new snippet' -PercentComplete (3/3 * 100)

        $ConfirmMsg = "`n`n`t$Activity`n`n"
        If ($PSCmdlet.ShouldProcess($Env:COMPUTERNAME,$ConfirmMsg)) {
        
            Try {
                
                # Update here-string            
                $SnippetBody = $Body.Replace("##AUTHOR##",$AuthorName)
                $SnippetBody = $Body.Replace("##CREATEDATE##",(Get-Date -f yyyy-MM-dd))
                $Param_Snip.Text = $SnippetBody

                # Create snippet
                $Create = New-IseSnippet @Param_Snip

                # Verify it
                If ($IsSuccess = Get-ISESnippet  -ErrorAction SilentlyContinue | 
                    Where-Object {($_.Name -eq $SnippetFile) -and ($_.CreationTime -lt (get-date))}) {
                    $Msg = "Snippet '$SnippetName' created successfully"
                    "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor Green
                    Write-Output $IsSuccess
                }
                Else {
                    $Msg = "Failed to create snippet '$SnippetName'"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[$Env:ComputerName] $Msg" | Write-MessageError
                }
            }
            Catch {
                $Msg = "Error creating snippet '$SnippetName'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                "[$Env:ComputerName] $Msg" | Write-MessageError
            }
        }
        Else {
            $Msg = "Snippet creation cancelled by user"
            "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor Cyan
        }
    }

}
End {
    Write-Progress -Activity $Activity -Completed

    $Msg = "END: $Activity"
    "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor Yellow -Title

}
} #end New-PKISESnippetFunctionInvokeCommand