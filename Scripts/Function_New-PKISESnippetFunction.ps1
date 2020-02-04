﻿#requires -Version 3
Function New-PKISESnippetFunction {
<#
.SYNOPSIS
    Adds a new PS ISE snippet containing a template function for ActiveDirectory, Invoke-Command, VMware, or generic use

.DESCRIPTION
    Adds a new PS ISE snippet containing a template function for ActiveDirectory, Invoke-Command, VMware, or generic use
    SupportsShouldProcess
    Returns a file object

.NOTES
    Name    : Function_New-PKISESnippetFunction.ps1
    Created : 2019-12-12
    Author  : Paula Kingsley
    Version : 01.0.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2019-12-12 - Created script based on older, separate snippet functions, now combined into one fun size
        
.PARAMETER Type
    Snippet content/type: ActiveDirectory, Generic, InvokeCommand, VMware

.PARAMETER AutoDetectAuthor
    Attempt to detect/construct current user's full name from registry

.PARAMETER Author
    Author name (if not using -AutoDetectAuthor)

.PARAMETER Force
    Overwrite existing snippet, if found

.EXAMPLE
    PS C:\> New-PKISESnippetFunction -Content ActiveDirectory -Verbose 

        VERBOSE: PSBoundParameters: 
	
        Key              Value                   
        ---              -----                   
        Type             ActiveDirectory         
        Verbose          True                    
        AutoDetectAuthor True                    
        Author                                   
        Force            False                   
        Quiet            False                   
        ScriptName       New-PKISESnippetFunction
        ScriptVersion    1.0.0                   

        VERBOSE: [Prerequisites] Automatically detected user's full name as 'Paula Kingsley'

        BEGIN: Create PowerShell ISE Snippet 'PK Active Directory Function'

        [LAPTOP12] Snippet 'PK Active Directory Function' created successfully

            Directory: C:\Users\paula\WindowsPowerShell\Snippets

        Mode                LastWriteTime         Length Name                                                                                                                                                                                                
        ----                -------------         ------ ----                                                                                                                                                                                                
        -a----       2019-12-12   3:17 PM          18322 PK Active Directory Function.snippets.ps1xml                                                                                                                                                        

        END  : Create PowerShell ISE Snippet 'PK Active Directory Function'

.EXAMPLE
    PS C:\> New-PKISESnippetFunction -Content InvokeCommand -Author "Barbara Hendricks" -Force -Quiet

        Directory: C:\Users\jbloggs\WindowsPowerShell\Snippets

        Mode                LastWriteTime         Length Name                                                                                                                                                                                                
        ----                -------------         ------ ----                                                                                                                                                                                                
        -a----       2019-12-12   3:19 PM          12826 PK Invoke-Command function.snippets.ps1xml       


#>
[Cmdletbinding(
    DefaultParameterSetName = "Auto",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    
    [Parameter(
        Mandatory = $True,
        HelpMessage = "Snippet content/type: ActiveDirectory, Generic, InvokeCommand, VMware"
    )]
    [ValidateSet("ActiveDirectory","Generic","InvokeCommand","VMware")]
    [Alias("Content")]
    [string]$Type,

    [Parameter(
        ParameterSetName = "Auto",
        HelpMessage = "Attempt to detect/construct current user's full name from registry"
    )]
    [switch]$AutoDetectAuthor,
    
    [Parameter(
        ParameterSetName = "Manual",
        HelpMessage = "Author name (if not using -AutoDetectAuthor)"
    )]
    [string]$Author,

    [Parameter(
        HelpMessage = "Overwrite existing snippet, if found"
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
    [version]$Version = "01.0.0000"
    
    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    If ($Source -eq "Auto") {$PSBoundParameters.AutoDetectAuthor = $AutoDetectAuthor = $True}
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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    # Function to get user's full name from registry based on WMI/SID
    Function GetFullName {
        $SIDLocalUsers = Get-WmiObject Win32_UserProfile -EA Stop | select-Object Localpath,SID
        $UserName = (Get-WMIObject -class Win32_ComputerSystem -Property UserName -ErrorAction Stop).UserName
        $UserOnly = $UserName.Split("\")[1]
        Foreach ($Profile in $SIDLocalUsers) {
            If ($Profile.localpath -like "*$UserOnly"){ # Match profile to current user
                    
                $SID = $Profile.sid # Look up path in registry/Group Policy by SID
                $DN = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\$SID" -ErrorAction SilentlyContinue | 
                    Select -ExpandProperty Distinguished-Name
                    
                If ($DN -match "(CN=)(.*?),.*") {    
                    $Results = ($DN -Split(",OU"))[0] -replace("CN=",$Null)
                        
                    # If the CN has a comma for the username, it probably means Last, First
                    If ($Results -match "\\,") {
                        $Results = $Results -replace("\\,",$Null)
                        $Surname = (($Results -split(" "))[0]).Trim()
                        $Given = ($Results -replace($Surname,$Null)).Trim()
                        $FullName = "$Given $Surname"
                    }
                    Else {$FullName = $Results}
                }
                If ($FullName) {Write-Output $FullName}
                Else {Write-Verbose "Failed to construct user's full name"}
            } #end if matching profile path
        } #end foreach
    } #end function

    #endregion Functions

    #region Prerequisites

    If (-not $PSISE) {
        $Msg = "This function requires the PowerShell ISE environment"
        "[$Env:ComputerName] $Msg" | Write-MessageError
        Break
    }

    Switch ($Source) {
        Manual {
            If ($CurrentParams.Author) {
                $Msg = "Author name manually specified as '$Author'"
                Write-Verbose "[Prerequisites] $Msg"
            }
            Else {
                $Msg = "Author name must be provided when -AuthorNameOption is set to 'Manual'"
                "[Prerequisites] $Msg" | Write-MessageError
                Break
            }
        }
        Auto   {
            If ($Author = GetFullName) {
                $Msg = "Automatically detected user's full name as '$Author'"
                Write-Verbose "[Prerequisites] $Msg"
            }
            Else {
                $Msg = "Failed to detect/construct current user's full name; please re-run script using the -Author parameter"
                "[Prerequisites] $Msg" | Write-MessageError
                Break
            }
        }
    }


    #endregion Prerequisites

    #region Snippet variables

    # Snippet name, description, and content (via here-string)

    Switch ($Type) {
    
        ActiveDirectory {
            $SnippetName = "PK Active Directory Function"
            $Description = "Snippet to create a new generic Active Directory function; created using New-PKISESnippetFunction -Type ActiveDirectory"
            
            $Body = @'
#Requires -version 4
Function Do-SomethingCool {
<# 
.SYNOPSIS
    Generic function using Active Directory module

.DESCRIPTION
    Generic function using Active Directory module
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
    One or more computernames (wildcards permitted)

.PARAMETER ADDomain             
    Active Directory domain name or FQDN (default is current user's domain)

.PARAMETER BaseDN               
    Starting Active Directory path/organizational unit (default is root of current user's domain)

.PARAMETER Server               
    Domain controller name or FQDN (default is first available)

.PARAMETER Quiet
    Hide all non-verbose console output

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
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage="One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

    [Parameter(
        HelpMessage = "Active Directory domain name or FQDN (default is current user's domain)"
    )]
    [ValidateNotNullOrEmpty()]
    [string] $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,

    [Parameter(
        HelpMessage = "Starting Active Directory path/organizational unit (default is root of current user's domain)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,
 
    [Parameter(
        HelpMessage = "Domain controller name or FQDN (default is first available)"
    )]
    [ValidateNotNullOrEmpty()]
    [String] $Server,

    [Parameter(
        HelpMessage = "Suppress non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch]$Quiet


)
Begin {
    
     # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

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

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"

    #region Functions
    # Mix, match, use, discard, whatever

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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    # Function to get the AD domain
    Function GetDomain {
        Param($ADDomain,$Credential)
        $Param_GetAD = @{}
        $Param_GetAD = @{
            Identity    = $ADDomain
            ErrorAction = "Stop"
            Verbose     = $False
        }
        If ($PSBoundParameters.Credential) {
            $Param_GetAD.Add("Credential",$Credential)
        }
        Try {
            Get-ADDomain @Param_GetAD
        }
        Catch {
            Throw $_.Exception.Message
        }
    }

    # Function to get the first availble DC in the domain
    Function GetFirstDC {
        Param($ADDomain)
        Try {
            $Param_GetDC = @{}
            $Param_GetDC = @{
                DomainName      = $ADDomain
                Discover        = $True
                NextClosestSite = $True
                ErrorAction     = "Stop"
                Verbose         = $False
            }
            Get-ADDomainController @Param_GetDC
        }
        Catch {
            Throw $_.Exception.Message
        }
    }

    # Function to get a named DC
    Function GetDC {
        Param($Server,$Credential)
        $Param_GetDC = @{}
        $Param_GetDC = @{
            Identity    = $Server
            Server      = $Server
            ErrorAction = "Stop"
            Verbose     = $False
        }
        If ($PSBoundParameters.Credential) {
            $Param_GetDC.Add("Credential",$Credential)
        }
        Try {
            Get-ADDomainController @Param_GetDC
        }
        Catch {
            Throw $_.Exception.Message
        }
    }

    # Function to verify searchbase
    Function GetSearchBase {
        Param($BaseDN,$Credential)
            $Param_GetSB = @{}
            $Param_GetSB = @{
                Identity    = $BaseDN
                Server      = $DC
                ErrorAction = "SilentlyContinue"
                Verbose     = $False
            }
        If ($PSBoundParameters.Credential) {
            $Param_GetSB.Add("Credential",$Credential)
        }
        Try {
            Get-ADObject @Param_GetSB
        }
        Catch {
            Throw $_.Exception.Message
        }
    }

    # Convert DN to CN
    function Get-CanonicalName ([string[]]$DistinguishedName) {    
        foreach ($dn in $DistinguishedName) {      
            $d = $dn.Split(',') ## Split the dn string up into it's constituent parts 
            $arr = (@(($d | Where-Object { $_ -notmatch 'DC=' }) | ForEach-Object { $_.Substring(3) }))  ## get parts excluding the parts relevant to the FQDN and trim off the dn syntax 
            [array]::Reverse($arr)  ## Flip the order of the array. 
 
            ## Create and return the string representation in canonical name format of the supplied DN 
            $("{0}/{1}" -f  (($d | Where-Object { $_ -match 'dc=' } | ForEach-Object { $_.Replace('DC=','') }) -join '.'), ($arr -join '/')).TrimEnd("/") 
        } 
    }

    #endregion Functions

    #region Prerequisites

    $Activity = "Prerequisites"

    # Make sure AD module is loaded
    # Use this only if not using a -Requires statement, which you may not want to do if this is part of a module containing non-AD-related functions

    $Msg = "Verify ActiveDirectory module"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    Try {
        If ($Module = Get-Module -Name ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue -Verbose:$False) {
            $Msg = "Successfully located ActiveDirectory module version $($Module.Version.ToString())"
            Write-Verbose "[Prerequisites] $Msg"
        }
        Else {
            $Msg = "Failed to find ActiveDirectory module in PSModule path"
            "[Prerequisites] $Msg" | Write-MessageError 
            Break
        }
    }
    Catch {
        $Msg = "Failed to find ActiveDirectory module"
        If ($ErrorDetails = $_.exception.message) {$Msg += " ($ErrorDetails)"}
        "[Prerequisites] $Msg" | Write-MessageError 
        Break
    }

    # Verify AD and a domain controller

    $Msg = "Connect to Active Directory"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    Try {
        $Param_GetAD = @{}
        $Param_GetAD = @{
            Identity    = $ADDomain
            ErrorAction = "Stop"
            Verbose     = $False
        }
        If ($CurrentParams.Credential) {
            $Param_GetAD.Add("Credential",$Credential)
        }
        $ADConfirm = Get-ADDomain @Param_GetAD
        $Msg = "Successfully connected to '$($ADConfirm.DNSRoot.Tolower())'"
        Write-Verbose "[Prerequisites] $Msg"
        
        # Get the domain controller
        If (-not $CurrentParams.Server) {
            
            $Msg = "Find nearest domain controller"
            Write-Verbose "[Prerequisites] $Msg"
            
            Try {        
                $Param_GetDC = @{}
                $Param_GetDC = @{
                    DomainName      = $ADDomain
                    Discover        = $True
                    NextClosestSite = $True
                    ErrorAction     = "Stop"
                    Verbose         = $False
                }
                $DCObj = Get-ADDomainController @Param_GetDC
                $DC = $($DCObj.HostName)
                $Msg = "Successfully connected to '$DC'"
                Write-Verbose "[Prerequisites] $Msg"
            }
            Catch {
                $Msg = "Failed to find domain controller for '$($ADConfirm.DNSRoot.Tolower())'"
                If ($ErrorDetails = $_.exception.message) {$Msg += "; $ErrorDetails"}
                "[Prerequisites] $Msg" | Write-MessageError
                Break
            }    
        }
        Else {
            $Msg = "Connect to named domain controller"
            Write-Verbose "[Prerequisites] $Msg"

            Try {        
                $Param_GetDC = @{}
                $Param_GetDC = @{
                    Identity    = $Server
                    ErrorAction = "Stop"
                    Verbose     = $False
                }
                If ($CurrentParams.Credential) {
                    $Param_GetDC.Add("Credential",$Credential)
                }
                $DCObj = Get-ADDomainController @Param_GetDC
                If ($DCObj.Domain -eq $ADConfirm.DNSRoot) {
                    $DC = $($DCObj.HostName)
                    $Msg = "Successfully connected to '$DC'"
                    Write-Verbose "[Prerequisites] $Msg"
                }
                Else {
                    $Msg = "Domain controller '$($DCObj.HostName)' is not in domain '$($ADConfirm.DNSRoot)'"
                    "[Prerequisites] $Msg" | Write-MessageError
                    Break
                }
            }
            Catch {
                $Msg = "Failed to find domain controller'$Server' in '$($ADConfirm.DNSRoot.Tolower())'"
                If ($ErrorDetails = $_.exception.message) {$Msg += "; $ErrorDetails"}
                "[Prerequisites] $Msg" | Write-MessageError
                Break
            }
        }    
        # Get or verify the searchbase
        If ($DCObj) {
                
            If ($CurrentParams.Searchbase) {

                $Msg = "Verify searchbase"
                Write-Verbose "[Prerequisites] $Msg"
                Write-Progress -Activity "Prerequisites" -CurrentOperation $Msg

                $Param_GetBaseDN = @{}
                $Param_GetBaseDN = @{
                    Identity    = $Searchbase
                    Server      = $DC
                    ErrorAction = "SilentlyContinue"
                    Verbose     = $False
                }

                Try {
                    # Overwriting the variable
                    If ([string]$Searchbase = Get-ADObject @Param_GetBaseDN | Select-Object -ExpandProperty DistinguishedName) {
                        $Separator = ""
                        $Msg = "Searchbase set to '$SearchBase'"
                        Write-Verbose "[Prerequisites] $Msg"
                    }
                    Else {
                        $Msg = "Failed to validate searchbase '$BaseDN' in domain $($ADConfirm.DNSRoot)'"
                        "[Prerequisites] $Msg" | Write-MessageError
                        Break
                    }
                }
                Catch {
                    $Msg = "Failed to validate searchbase '$BaseDN'"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[Prerequisites] $Msg" | Write-MessageError
                    Break
                }
            }    
            Else {
                $Msg = "Set searchbase"
                Write-Verbose "[Prerequisites] $Msg"
                Write-Progress -Activity "Prerequisites" -CurrentOperation $Msg

                [string]$Searchbase = $ADConfirm.DistinguishedName
                $Msg = "Successfully created searchbase as '$SearchBase'"
                Write-Verbose "[Prerequisites] $Msg"
            } 
        } 
    }
    Catch [exception] {
        $Msg = "Failed to connect to AD domain '$ADDomain'"
        If ($ErrorDetails = $_.exception.Message) {$Msg += " ($ErrorDetails)"}
        "[Prerequisites] $Msg" | Write-MessageError
        Break
    }

    #endregion Prerequisites

    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
        Debug       = $False
    }

    # Splat for Get-ADComputer
    $Props = "Name","IPv4Address","Description","Location","OperatingSystem","DistinguishedName","CanonicalName","SID","WhenCreated","ServicePrincipalNames"
    $Select = "Name","IPv4Address","Description","Location","OperatingSystem","DistinguishedName","CanonicalName","SID","WhenCreated","ServicePrincipalNames",@{N="OU";E={$_.DistinguishedName -replace '^.+?(?<!\\),',''}}
    $Param_AD = @{}
    $Param_AD = @{
        Filter      = $Null
        Properties  = $Props
        Searchbase  = $SearchBase
        SearchScope = "Subtree"
        Server      = $DC
        ErrorAction = "Stop"
        Verbose     = $False
        Debug       = $False
    }

    # Splat for write-progress
    $Activity = "Do AD stuff"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        PercentComplete  = $Null
        CurrentOperation = $Null
        Status           = "Working"
    }

    #endregion Splats

    #region Functions

    # Convert DN to CN
    function Get-CanonicalName ([string[]]$DistinguishedName) {    
        foreach ($dn in $DistinguishedName) {      
            $d = $dn.Split(',') ## Split the dn string up into it's constituent parts 
            $arr = (@(($d | Where-Object { $_ -notmatch 'DC=' }) | ForEach-Object { $_.Substring(3) }))  ## get parts excluding the parts relevant to the FQDN and trim off the dn syntax 
            [array]::Reverse($arr)  ## Flip the order of the array. 
 
            ## Create and return the string representation in canonical name format of the supplied DN 
            $("{0}/{1}" -f  (($d | Where-Object { $_ -match 'dc=' } | ForEach-Object { $_.Replace('DC=','') }) -join '.'), ($arr -join '/')).TrimEnd("/") 
        } 
    }

    #endregion Functions  

    # Console output
    
    "BEGIN: $Activity" | Write-MessageInfo -FGColor -Yellow -Title
    

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
            
        $ConfirmMsg = "`n`n`t$Activity`n`n"
        If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
            Try {
                $Msg = $Activity
                $Param_WP.CurrentOperation = $Msg
                Write-Verbose "[$Computer] $Msg" 
                Write-Progress @Param_WP

                If ($Computer -match "\*") {
                    $Param_AD.Filter =  "Name -like '$Computer'"
                }
                Else {
                    $Param_AD.Filter =  "Name -eq '$Computer'"
                }
                Get-ADComputer @Param_AD 
                
            }
            Catch {
                $Msg = "Operation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                "[$Computer] $Msg" | Write-MessageError
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
    "END  : $Activity" | Write-MessageInfo -FGColor -Yellow -Title
}

} # end Do-SomethingCool


'@

        } #end ActiveDirectory
        Generic {
            
            $SnippetName = "PK Generic Function"
            $Description = "Snippet to create a new generic function; created using New-PKISESnippetFunction -Type Generic"
        
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
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
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
        HelpMessage="Hide all non-verbose console output"
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
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    # Output
    [array]$Results = @()
    
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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
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
    $Activity = "Do a thing"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }

    # Splat for whatever
    $Param_Thing = @{}
    $Param_Thing = @{
        ComputerName = $Null
        Verbose      = $False
        ErrorAction  = "Stop"
    }

    #endregion Splats

    # Console output
    "BEGIN: $Activity" | Write-MessageInfo -FGColor Yellow -Title
    


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

        $Msg = "Do a thing"
        "[$Computer] $Msg" | Write-MessageInfo -FGColor White
        
        If ($Continue.IsPresent) {
            
            $ConfirmMsg = "`n`n`t$Activity`n`n"
            If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
                Try {
                    $Msg = $Activity
                    $Param_WP.CurrentOperation = $Msg
                    

                    Write-Progress @Param_WP

                    $Param_Thing.ComputerName = $Computer
                    Do-Things

                    $Msg = "Did a thing"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor Green
                    
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[$Computer] $Msg" | Write-MessageError
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor White
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    "END  : $Activity" | Write-MessageInfo -FGColor Yellow -Title
}

} # end Do-SomethingCool


'@
    
        } #end Generic
        InvokeCommand {
            $SnippetName = "PK Invoke-Command function"
            $Description = "Snippet to create a new generic function to run Invoke-Command; created using New-PKISESnippetFunction -Type VMware"

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

.PARAMETER Authentication
    WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'Job')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests ignored on local computer)

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
        HelpMessage = "WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)"
    )]
    [ValidateSet('Kerberos','Basic','Negotiate','Default','CredSSP')]
    [string]$Authentication = "Negotiate",

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
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests are ignored on local computer)"
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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    # Function to test WinRM connectivity
    Function Test-WinRM{
        Param($Computer)
        $Param_WSMAN = @{
            ComputerName   = $Computer
            Credential     = $Credential
            Authentication = $Authentication
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
        Authentication = $Authentication
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
        [bool]$IsLocal = $True

        If ($Computer -in ($Env:ComputerName,"localhost","127.0.0.1")) {
            $IsLocal = $True
            $Continue = $True
        }
        Else {
            Switch ($ConnectionTest) {
                Default {$Continue = $True}
                Ping {
                    $Msg = "Ping computer"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
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
                WinRM {
                    $Msg = "Test WinRM connection"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
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
            }
        } #end if not local

        If ($Continue.IsPresent) {
            
            $ConfirmMsg = "`n`n`t$Activity`n`n"
            If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
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
    
        } # end InvokeCommand
        VMware {
            $SnippetName = "PK VMware Function"
            $Description = "Snippet to create a new generic VMware/PowerCLI function; created using New-PKISESnippetFunction -Type VMware"

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

.PARAMETER Quiet
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
        Mandatory = $True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more VMs or VM names"
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
        HelpMessage = "Suppress non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch]$Quiet

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
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

    # Function to write an error as a string (no stacktrace)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$Force)
        $Host.UI.WriteErrorLine("$Message")
    }

    # Function to test vCenter connection
    Function TestVISession {
    [CmdletBinding()]
    Param([string]$Server,[switch]$BooleanOutput,[switch]$Quiet)
        $VCenter = $Null
        If ($PSBoundParameters.Server) {
            Try { 
                If ($vCenter = Get-Variable DefaultVIServers -Scope Global -ErrorAction SilentlyContinue | Where-Object {$_.Value.Name -Match $Server}) {
                    $Msg = "Connection found to vCenter server '$Server'"
                    If ($BooleanOutput.IsPresent) {$True}
                    Else {Write-Output $vCenter.Value.Name}
                }
                Else {
                    $Msg = "No connection found to vCenter server '$Server'"
                    If ($BooleanOutput.IsPresent) {$False}   
                    Else {Write-Output $Msg}
                }
            }
            Catch {}
        }
        Else {
            Try {
                If ($vCenter = Get-Variable DefaultVIServers -Scope Global -ErrorAction SilentlyContinue) {
                    $Msg = "Connection found to vCenter"
                    If ($BooleanOutput.IsPresent) {$True}
                    Else {Write-Output $vCenter.Value}
                }
                Else {
                    $Msg = "No connection found to vCenter"
                    If ($BooleanOutput.IsPresent) {$False}
                    Else {Write-Output $Msg}
                }
            }
            Catch {}
        }
    } #end TestViSession

    #endregion Functions

    #region Prerequisites
    
    $Msg = "Validate PowerCLI module"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Msg

    # Check for PowerCLI...using this instead of #Requires to ensure entire GNOpsWindowsVM module loads, regardless of PowerCLI availability
    Try {
        If ($PCLIMod = Get-Module VMware.PowerCLI -ErrorAction SilentlyContinue -Verbose:$False -Debug:$False) {
            $Msg = "Verified PowerCLI module is available (version $($PCLIMod.Version.toString()))"
            Write-Verbose "[Prerequisites] $Msg"
        }
        Else {
            If ($PCLIMod = Get-Module VMware.PowerCLI -ErrorAction SilentlyContinue -Verbose:$False -Debug:$False -ListAvailable | Import-Module -Force -PassThru -Verbose:$False -ErrorAction -Stop) {
                $Msg = "Verified PowerCLI module is available (version $($PCLIMod.Version.toString()))"
                Write-Verbose "[Prerequisites] $Msg"
            }
            Else {
                $Msg = "Failed to detect VMware PowerCLI module loaded in this session; please ensure it is already loaded or install it from https://www.powershellgallery.com/packages/VMware.PowerCLI"
                "[Prerequisites] $Msg" | Write-MessageError -Force
                Break
            }
        }
    }
    Catch {
        $Msg = "Failed to detect VMware PowerCLI module loaded in this session; please ensure it is installed from https://www.powershellgallery.com/packages/VMware.PowerCLI"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n $ErrorDetails"}
        "[Prerequisites] $Msg" | Write-MessageError -Force
        Break
    }

    $Msg = "Test vCenter connection"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Msg
    
    If ($CurrentParams.VIServer) {
        
        $VIServer | Foreach-Object {
            If (TestVISession -Server $_ -Verbose:$False) {
                $Msg = "Connection found to vCenter '$_'"
                Write-Verbose "[Prerequisites] $Msg"
            }
            Else {
                $Msg = "No connection found to vCenter '$_'"
                "[Prerequisites] $Msg" | Write-MessageError -Force
                Break
            }
        }
    }
    Else {
        If ($VIServer = TestVISession -Verbose:$False) {
            $Msg = "Connection found to vCenter '$($VIServer -join(''', '''))'"
            Write-Verbose "[Prerequisites] $Msg"
        }
        Else {
            $Msg = "No connection found to vCenter"
            "[Prerequisites] $Msg" | Write-MessageError -Force
            Break
        }
    }
    $VIServerStr = "'$($VIServer -join(''', '''))'"

    If ($VIServer) {
        # Set the timeout to something high
        Try {
            $SetTimeout = Set-PowerCLIConfiguration -WebOperationTimeoutSeconds 3600 -Scope Session -Verbose:$False -Debug:$False -Confirm:$False -ErrorAction SilentlyContinue
        }
        Catch {
            $Msg = "Failed to set PowerCLI WebOperationTimeoutSeconds to 3600; lengthy operations may fail"
            If ($ErrorDetails = $_.exception.Message) {$Msg += " ($ErrorDetails)"}
            Write-Warning $Msg
        }
    }   

    #endregion Prerequisites

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
        Name        = $Null
        Server      = $VIServer
        ErrorAction = "Stop"
        Verbose     = $False
    }
    
    #endregion Splats
 
    # Console output
    "BEGIN  : $Activity" | Write-MessageInfo -FGColor Yellow -Title

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
                Write-Progress @Param_WP

                "[$V] $Msg" | Write-MessageInfo -FGColor White

                $Param_VM.Name = $V
                $VMObj = Get-VM @Param_VM

                $Msg = "Successfully got VM"
                "[$($VMObj.Name)] $Msg" | Write-MessageInfo -FGColor Green

                Write-Output $VMObj
            }
            Catch {
                $Msg = "Operation failed"
                If ($ErrorDetails = $($_.exception.message -replace '\s+', ' ')) {$Msg += " ($ErrorDetails)"}
                "[$V] $Msg" | Write-MessageError
            }
        }
        Else {
            $Msg = "Operation cancelled by user"
            "[$V] $Msg" | Write-MessageInfo -FGColor White
        }
        
    } #end for each computer
        
}
End {
    
    Write-Progress -Activity $Activity -Completed
    "END    : $Activity" | Write-MessageInfo -FGColor Yellow -Title
}

} # end Do-SomethingCool


'@

        } # end VMware
    
    } #end switch for content 

    # Construct filename
    $SnippetFile = $SnippetName +".snippets.ps1xml"

    #endregion Snippet variables

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
    $Activity = "Create PowerShell ISE Snippet '$SnippetName'"
    "BEGIN: $Activity"| Write-MessageInfo -FGColor Yellow -Title
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
    Else {
        
        Write-Progress -Activity $Activity -Status $Env:ComputerName -CurrentOperation 'Create new snippet' -PercentComplete (3/3 * 100)

        $ConfirmMsg = "`n`n`t$Activity`n`n"
        If ($PSCmdlet.ShouldProcess($Env:COMPUTERNAME,$ConfirmMsg)) {
        
            Try {
                
                # Update here-string            
                $SnippetBody = $Body.Replace("##AUTHOR##",$Author)
                $SnippetBody = $SnippetBody.Replace("##CREATEDATE##",(Get-Date -f yyyy-MM-dd))
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
    "END  : $Activity" | Write-MessageInfo -FGColor Yellow -Title

}
} #end New-PKISESnippetFunction