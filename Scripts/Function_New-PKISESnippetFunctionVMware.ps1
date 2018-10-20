#requires -Version 4
Function New-PKISESnippetFunctionVMware {
<#
.SYNOPSIS
    Adds a new PS ISE snippet containing a template function using the PowerCLI module

.DESCRIPTION
    Adds a new PS ISE snippet containing a template function using the PowerCLI module
    SupportsShouldProcess
    Returns a file object

.NOTES
    Name    : Function_New-PKISESnippetFunctionVMware.ps1
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
    PS C:\> New-PKISESnippetFunctionVMware -Author foo -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                             
        ---                      -----                             
        Author                   foo                               
        Verbose                  True                              
        AutoDetectAuthorFullName False                             
        Force                    False                             
        Confirm                  PK VMware Function.snippets.ps1xml
        ScriptName               New-PKISESnippetFunctionVMware    
        ScriptVersion            1.0.0                             

        VERBOSE: Setting author name to 'foo'
        Action: Create ISE Snippet 'PK VMware Function'
        Snippet creation cancelled by user

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionVMware -AutoDetectAuthorFullName -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                             
        ---                      -----                             
        AutoDetectAuthorFullName True                              
        Force                    False
        Verbose                  True                              
        Author                                                     
        Confirm                  PK VMware Function.snippets.ps1xml
        ScriptName               New-PKISESnippetFunctionVMware    
        ScriptVersion            1.0.0                             

        VERBOSE: Setting author to current user's full name, 'Paula Kingsley'
        Action: Create ISE Snippet 'PK VMware Function'
        VERBOSE: Snippet 'PK VMware Function' created successfully

            Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                                        
        ----                -------------         ------ ----                                                                                        
        -a----       2018-10-19  05:29 PM          23925 PK VMware Function.snippets.ps1xml

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionVMware -AutoDetectAuthorFullName
    
        Action: Create ISE Snippet 'PK VMware Function'
        Snippet 'PK VMware Function' already exists; specify -Force to overwrite

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

    $SnippetName = "PK VMware Function"
    $Description = "Snippet to create a new generic VMware/PowerCLI function; created using New-PKISESnippetFunctionVMware"

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
#Requires -version 4
Function Do-SomethingCool {
<# 
.SYNOPSIS
    Generic function using PowerCLI module

.DESCRIPTION
    Generic function using PowerCLI module
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
        
.PARAMETER VM        
    One or more VM names or objects

.PARAMETER VIServer
    One or more VCenter servers (default is any connected)

.PARAMETER ModuleVersion               
    Minimum required PowerCLI version

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Do-SomethingCool -VM foo

        
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
        HelpMessage="One or more VMs or VM names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","ComputerName","Name","Guest")]
    [Object[]] $VM,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "One or more VCenter servers (default is any connected)"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$VIServer,

    [Parameter(
        ParameterSetName = "Version",
        HelpMessage = "Minimum required PowerCLI module version"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$ModuleVersion = "6.5.4",

    [Parameter(
        ParameterSetName = "Version",
        HelpMessage = "Enforce minimum required PowerCLI module version"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$EnforceMinimumVersion,

    [Parameter(
        HelpMessage = "Suppress non-verbose, non-error console output"
    )]
    [Switch]$SuppressConsoleOutput

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("VM")) -and (-not $VM)
    $Source = $PSCmdlet.ParameterSetName

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    #region Prerequisites 
    $Activity = "Prerequisites"

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"

    #region Functions

    # Inner function to test PowerCLI
    Function TestPCLI {
        [CmdletBinding()]
        Param(
            $Major = "0",
            $Minor = "0",
            $Build = "0",
            $Revision = "0",
            [switch]$EnforceVersion,
            [switch]$WarnVersion,
            [switch]$IgnoreVersion = $True
        )
        Process {
            # We can't warn or enforce *and* ignore....
            If ($EnforceVersion.IsPresent -or $WarnVersion.IsPresent) {$IgnoreVersion = $False}
            Try {
                $ModuleObj = $Null
                If (-not ($ModuleObj = Get-Module -Name VMware.PowerCLI -ErrorAction SilentlyContinue -Verbose:$False)) {
                    $Msg = "Required module not found; please see https://blogs.vmware.com/PowerCLI/2017/04/powercli-install-process-powershell-gallery.html"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    Break
                }
                If ($IgnoreVersion.IsPresent) {
                    $Msg = "PowerCLI module $($ModuleObj.Version.ToString()) detected"
                    Write-Verbose "$Msg"
                    $True
                }
                Else {
                    [version]$MinVer = ('{0}.{1}.{2}.{3}' -f $Major,$Minor,$Build,$Revision)
                    If (($CurrModVer = ($ModuleObj.Version)) -lt $Minver) {
                        If ($EnforceVersion.IsPresent) {
                            $Msg = "This script requires VMware PowerCLI version $($Minver.ToString()) or greater`nCurrent version detected is $($ModuleObj.Version.ToString())"
                            $Host.UI.WriteErrorLine($Msg)
                            $False
                        }
                        Elseif ($WarnVersion.IsPresent) {
                            $Msg = "This script requires VMware PowerCLI version $($Minver.ToString()) or greater`nCurrent version detected is $($ModuleObj.Version.ToString()); you may encounter errors`nPlease consider upgrading using Chocolatey or the PSGallery"
                            Write-Warning $Msg
                            $False
                        }
                    }
                    Else {
                        $Msg = "PowerCLI module $($ModuleObj.Version.ToString()) detected (minimum required version is $($Minver.ToString()))"
                        Write-Verbose "$Msg"
                        $True
                    }
                }
            }
            Catch {
                $Msg = "Failed to locate required module 'VMware.PowerCLI'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
            }
        }
    }

    # Inner function test VI session
    Function TestViSession {
    [CmdletBinding(DefaultParameterSetName = "All")]
    Param( 

        [Parameter(
            ParameterSetName = "Match",
            Mandatory=$True,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            Position = 0,
            HelpMessage="Name of VCenter Server to test connection"
        )]
        [Parameter(
            ParameterSetName = "All",
            Mandatory=$False,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            Position = 0,
            HelpMessage="Name of VCenter Server to test connection"
        )]
        [ValidateNotNullOrEmpty()]
        [Alias("Name")]
        [string[]] $VIServer,

        [Parameter(
            ParameterSetName = "Match",
            HelpMessage="Test connection to named server"
        )]
        [Switch]$MatchServerName,

        [Parameter(
            HelpMessage="Return Boolean or PSObject details"
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Full","Boolean")]
        [String]$OutputType = "Full"

    )
    Begin {
        $Source = $PSCmdlet.ParameterSetName

        # Current version (please keep up to date from comment block)
        [version]$Version = "03.03.0000"

        # Show our settings
        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
            Where {Test-Path variable:$_}| Foreach {
                $CurrentParams.Add($_, (Get-Variable $_).value)
            }
        $CurrentParams.Add("ParameterSetName",$Source)
        $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
        $CurrentParams.Add("ScriptVersion",$Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        # Preference
        $ErrorActionPreference = "Stop"
        $ProgressPreference = "Continue"
    }
    Process {
    
        Try {

            Switch ($Source) {
                All {
                    $Msg = "Get all connected session(s)"
                    Write-Verbose $Msg
                    $Sessions = (Get-View 'SessionManager' -ErrorAction SilentlyContinue -Verbose:$False -Debug:$False)
            
                    [array]$AllOutput = @()
                    Foreach ($Session in $Sessions) {
                        $Output = New-Object PSObject -Property ([ordered] @{
                            VIServer     = ($Session.Client.ServiceUrl -as [System.Uri]).Host
                            ServiceURL   = $Session.Client.ServiceUrl
                            UserName     = $Session.CurrentSession.UserName
                            FullName     = $Session.CurrentSession.FullName
                            LoginTime    = $Session.CurrentSession.LoginTime
                            IPAddress    = $Session.CurrentSession.IPAddress
                            UserAgent    = "$($Session.CurrentSession.UserAgent)"
                            ComputerName = $Env:ComputerName
                            Messages     = $Null
                        })
                        $AllOutput += $Output
                    }
                } # end all
                Match {
                    $Msg = "Look for connected sessions to server(s) '$($VIServer -join("', '"))'"
                    Write-Verbose $Msg
                    [array]$Session = @()
                    $Sessions = ((Get-View 'SessionManager' -ErrorAction SilentlyContinue -Verbose:$False -Debug:$False) | 
                        Where-Object {(($_.Client.ServiceUrl -as [System.Uri]).Host) -match $VIServer})   
            
                    [array]$AllOutput = @()
                    Foreach ($Session in $Sessions) {
                        $Output = New-Object PSObject -Property ([ordered] @{
                            VIServer     = ($Session.Client.ServiceUrl -as [System.Uri]).Host
                            ServiceURL   = $Session.Client.ServiceUrl
                            UserName     = $Session.CurrentSession.UserName
                            FullName     = $Session.CurrentSession.FullName
                            LoginTime    = $Session.CurrentSession.LoginTime
                            IPAddress    = $Session.CurrentSession.IPAddress
                            UserAgent    = "$($Session.CurrentSession.UserAgent)"
                            ComputerName = $Env:ComputerName
                            Messages     = $Null
                        })
                        $AllOutput += $Output
                    }
                }
            } # end switch for parametersetname

            If ($AllOutput.Count -gt 0) {
        
                Switch ($Source) {
                    All {
                        $Msg = "{0} current VCenter connection(s) found" -f $AllOutput.VIServer.Count
                    }
                    Match {
                        $Msg = "{0} current VCenter connection(s) found to '$($VIServer -join("', '"))'" -f $AllOutput.VIServer.Count
                    }
                }
                Write-Verbose $Msg
                Switch ($OutputType) {
                    Boolean {
                        $True
                    }
                    Full {
                        Write-Output $AllOutput            
                    }
                }
            }
            Else {
                Switch ($Source) {
                    All {
                        $Msg = "No current VCenter connection(s) found"
                    }
                    Match {
                        $Msg = "No current VCenter connection(s) found to '$($VIServer -join("', '"))'"
                    }
                }
                Write-Warning $Msg
                Switch ($OutputType) {
                    Boolean {
                        $False
                    }
                    Full {
                        New-Object PSObject -Property ([ordered] @{
                            VIServer     = "-"
                            ServiceURL   = "-"
                            UserName     = "-"
                            FullName     = "-"
                            LoginTime    = "-"
                            IPAddress    = "-"
                            UserAgent    = "-"
                            ComputerName = $Env:ComputerName
                            Messages     = $Msg
                        })          
                    }
                }
            } # end if no results
        }
        Catch {
            $Msg = "No VCenter connections found"
            If ($ErrorDetails = "$($_.exception.message -replace '\s+', ' ')") {$Msg += "; $ErrorDetails"}
            Write-Warning "$Msg"
            Switch ($OutputType) {
                "Full" {
                    New-Object PSObject -Property ([ordered] @{
                        VIServer     = $Null
                        ServiceURL   = $Null
                        UserName     = $Null
                        FullName     = $Null
                        LoginTime    = $Null
                        IPAddress    = $Null
                        UserAgent    = $Null
                        ComputerName = $Env:ComputerName
                        Messages     = $Msg
                    })
                }
                "Boolean" {
                    $False
                }
            }
        }
        Finally {
            Write-Progress -Activity $Activity -Completed -Verbose:$False -ErrorAction SilentlyContinue -Debug:$False
        }
    } #end process
    End {}
    } #end TestVISession

    #endregion Functions

    #region Prerequisites

    $Param_TestSession = @{}
    $Param_TestSession = @{
        Verbose               = $False
        ErrorAction           = "SilentlyContinue"
    }
    If ($CurrentParams.VIServer) {
        $Param_TestSession.Add("VIServer",$VIServer)
        $Param_TestSession.Add("MatchServer",$True)
        
        Try {
            If (($Null = Test-ViSession @Param_TestSession) -eq $True)  {
                $Msg = "Verified connection to '$VIServer'"
                Write-Verbose "[Prerequisites] $Msg"

            }
            Else {
                $Msg = "No current connection found to VCenter server '$VIServer'; script will now exit"
                $Host.UI.WriteErrorLine("ERROR: $Msg") 
                Break
            }
        }
        Catch {
            $Msg = "VCenter connection test failed"
            If ($ErrorDetails = $_.exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
    }
    Else {
        Try {

            If (($VIServer = (TestViSession @Param_TestSession).VIServer))  {
                
                $Msg = "Verified connection(s) to '$($VIServer -join("', '"))'"
                Write-Verbose "[Prerequisites] $Msg"
            }
            Else {
                $Msg = "No current connection found to VCenter; script will now exit"
                $Host.UI.WriteErrorLine("ERROR: $Msg") 
                Break
            }
        }
        Catch {
            $Msg = "VCenter connection test failed"
            If ($ErrorDetails = $_.exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
    
    } 
    
    If ($VIServer) {
        # Set the timeout to something high
        Try {
            $SetTimeout = Set-PowerCLIConfiguration -WebOperationTimeoutSeconds 3600 -Scope Session -Verbose:$False -Debug:$False -Confirm:$False -ErrorAction SilentlyContinue
        }
        Catch {
            $Msg = "Failed to set PowerCLI WebOperationTimeoutSeconds to 3600; lengthy operations may fail"
            If ($ErrorDetails = $_.exception.Message) {$Msg += "`n$ErrorDetails"}
            Write-Warning $Msg
        }
    }   
    
    # Test PowerCLI
    # For outer progress bar
    $Activity = "Prereqisites"
    $Msg = "Test PowerCLI module availability and version"
    $CurrentOp = $Msg
    Write-Progress -Activity $Activity -CurrentOp $Msg

    Try {
        Write-Verbose "$Msg"
        $Param_PCLI = @{}
        $Param_PCLI = @{
            Verbose     = $False
            ErrorAction = "Stop"
        }
        If ($EnforceMinimumVersion.IsPresent) {
            $VersionObj = ($ModuleVersion -as [Version])
            $Param_PCLI.Add("Major",$VersionObj.Major)
            $Param_PCLI.Add("Minor",$VersionObj.Minor)
            $Param_PCLI.Add("Build",$VersionObj.Build)
            $Param_PCLI.Add("EnforceVersion",$True)
            
            If ($Null = (TestPCLI @Param_PCLI)) {
                $Msg = "Minimum or greater version of PowerCLI module found"
                Write-Verbose $Msg
            }
            Else {
                $Msg = "Minimum or greater version of PowerCLI module not found"
                $Host.UI.WriteErrorLine("ERROR: $Msg") 
                Break
            }
        }
        Else {
            If ($Null = (TestPCLI @Param_PCLI)) {
                $Msg = "PowerCLI module found"
                Write-Verbose $Msg
            }
            Else {
                $Msg = "PowerCLI module not found"
                $Host.UI.WriteErrorLine("ERROR: $Msg") 
                Break
            }
        }
    }
    Catch {
        $Msg = "Failed to verify PowerCLI module"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
        Break
    }

    # Test Vcenter connection
    # For outer progress bar
    $Activity = "Prereqisites"
    $Msg = "Test Vcenter connection"
    $CurrentOp = $Msg
    Write-Progress -Activity $Activity -CurrentOp $Msg

    Try {
        Write-Verbose "$Msg"
        $Param_VC = @{}
        $Param_VC = @{
            Verbose = $False
            ErrorAction = "Stop"
        }
        If ($PSBoundParameters.ContainsKey("VIServer")) {
            
            $Param_VC.Add("VIServer",$VIServer)
            $Param_VC.Add("MatchServerName",$True)
            
            If ($Null = (TestViSession @Param_VC)) {
                $Msg = "Verified connection to Vcenter '$VIServer'"
                Write-Verbose $Msg
            }
            Else {
                $Msg = "Failed to verify connection to Vcenter '$VIServer'"
                $Host.UI.WriteErrorLine("ERROR: $Msg") 
                Break
            }
        }
        Else {
            If ($VIServer = (TestViSession @Param_VC -OutputType Full).VIServer) {
                $Msg = "Verified connection to '$($VIServer -join("','"))'"
                Write-Verbose $Msg
            }
            Else {
                $Msg = "No connection found to Vcenter"
                $Host.UI.WriteErrorLine("ERROR: $Msg") 
                Break
            }
        }
    }
    Catch {
        $Msg = "Failed to confirm Vcenter connection"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
        Break
    }

    $Msg = "Verified connection to '$VIServer'"
    Write-Verbose "[Prerequisites] $Msg"

    # Set the timeout to something high
    Try {
        $SetTimeout = Set-PowerCLIConfiguration -WebOperationTimeoutSeconds 3600 -Scope Session -Verbose:$False -Debug:$False -Confirm:$False -ErrorAction SilentlyContinue
    }
    Catch {
        $Msg = "Failed to set PowerCLI WebOperationTimeoutSeconds to 3600; lengthy operations may fail"
        If ($ErrorDetails = $_.exception.Message) {$Msg += "`n$ErrorDetails"}
        Write-Warning $Msg
    }
      
    #endregion Prerequisites

    #region Output object and property order

    # Will be returned if no match found
    $InitialValue = "Error"
    $OutputTemplate = New-Object PSObject -Property ([ordered]@{
        Computername = $InitialValue        
        Messages     = $InitialValue
    })
    $Select = $OutputTemplate.PSObject.Properties.Name

    #endregion Output object and property order
    
    #region Splats

    # General-purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
        Debug       = $False
    }

    # Splat for Write-Progress
    $Activity = "Generic VMware function"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        PercentComplete  = $Null
        CurrentOperation = $Null
        Status           = "Working"
    }

    # Splat for Get-VM
    $Param_VM = @{}
    $Param_VM = @{
        Name = $Null
        Server = $VIServer
        ErrorAction = "Stop"
        Verbose = $False
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
    $Total = $VM.Count
    $Current = 0

    $Results = @()

    Foreach ($V in $VM) {
        
        $V = $V.Trim()
        $Current ++ 

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = $V
            
        $ConfirmMsg = "`n`n`t$Activity`n`n"
        If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
            Try {
                $Msg = "Get VM"
                $Param_WP.CurrentOperation = $Msg
                Write-Verbose "[$V] $Msg"
                Write-Progress @Param_WP

                $Param_VM.Name = $V
                $VMObj = Get-VM @Param_VM
                $Results += $VMObj
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
        Author      = $AuthorName
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
} #end New-PKISESnippetFunctionVMware