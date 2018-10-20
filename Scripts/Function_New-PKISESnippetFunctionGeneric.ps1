#requires -Version 3
Function New-PKISESnippetFunctionGeneric {
<#
.SYNOPSIS
    Adds a new PS ISE snippet containing a template function

.DESCRIPTION
    Adds a new PS ISE snippet containing a template function
    SupportsShouldProcess
    Returns a file object

.NOTES
    Name    : Function_New-PKISESnippetFunctionGeneric.ps1
    Created : 2018-10-15
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2018-10-15 - Created script

.PARAMETER Author
    Author name

.PARAMETER AutoDetectAuthorFullName
    Attempt to match the current username to their full name via the registry & WMI
    .PARAMETER Force
    Forces creation even if snippet name exists

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionGeneric -Author "Diana Prince" -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                          
        ---                      -----                          
        Author                   Diana Prince                   
        Verbose                  True                           
        AutoDetectAuthorFullName False                          
        Force                    False                          
        Confirm                                                 
        ScriptName               New-PKISESnippetFunctionGeneric
        ScriptVersion            1.0.0                          



        VERBOSE: Setting author name to 'Diana Prince'
        Action: Create ISE Snippet 'PK Generic Function'
        VERBOSE: Snippet 'PK Generic Function' created successfully


            Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                                        
        ----                -------------         ------ ----                                                                                        
        -a----       2018-10-19  05:49 PM           8710 PK Generic Function.snippets.ps1xml      

.EXAMPLE
    New-PKISESnippetFunctionGeneric -AutoDetectAuthorFullName -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                          
        ---                      -----                          
        AutoDetectAuthorFullName True                           
        Verbose                  True                           
        Author                                                  
        Force                    False                          
        Confirm                                                 
        ScriptName               New-PKISESnippetFunctionGeneric
        ScriptVersion            1.0.0                          

        VERBOSE: Setting author to current user's full name, 'Paula Kingsley'
        Action: Create ISE Snippet 'PK Generic Function'
        Snippet 'PK Generic Function' already exists; specify -Force to overwrite
        VERBOSE: 

            Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                                        
        ----                -------------         ------ ----                                                                                        
        -a----       2018-10-19  05:49 PM           8710 PK Generic Function.snippets.ps1xml                                                         

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionGeneric -AutoDetectAuthorFullName -Force -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                          
        ---                      -----                          
        AutoDetectAuthorFullName True                           
        Force                    True                           
        Verbose                  True                           
        Author                                                  
        Confirm                                                 
        ScriptName               New-PKISESnippetFunctionGeneric
        ScriptVersion            1.0.0                          

        VERBOSE: Setting author to current user's full name, 'Paula Kingsley'
        Action: Create ISE Snippet 'PK Generic Function'
        VERBOSE: Snippet 'PK Generic Function' created successfully


            Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                                        
        ----                -------------         ------ ----                                                                                        
        -a----       2018-10-19  05:49 PM           8700 PK Generic Function.snippets.ps1xml  

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
        $Msg = "This function requires the PowerShell ISE environment"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }

    #region Snippet info
    
    $SnippetName = "PK Generic Function"
    $Description = "Snippet to create a new generic function; created using New-PKISESnippetFunctionGeneric"

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
    Generic function

.DESCRIPTION
    Generic function
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Do-SomethingCool.ps1
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
    $Activity = "Run command"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
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
            
            $ConfirmMsg = "`n`n`t$Activity`n`n"
            If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
                Try {
                    $Msg = $Activity
                    $Param_WP.CurrentOperation = $Msg
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress @Param_WP

                    $Param_IC.ComputerName = $Computer
                    $Results += Do Things @Param_IC
                    
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

     If ($Results.Count -eq 0) {
        $Msg = "No output returned"
        Write-Warning $Msg
    }
    Else {
        $Msg = "$($Results.Count) result(s) found"
        Write-Verbose $Msg
        Write-Output ($Results | Select -Property *)
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
} #end New-PKISESnippetFunctionGeneric