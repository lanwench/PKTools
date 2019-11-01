#Requires -version 3
Function Set-PKWMIPermissions {
<# 
.SYNOPSIS
    Invokes a scriptblock to set WMI permissions, interactively or as a PSJob, with a huge HT to Graeme Bray & Steve Lee

.DESCRIPTION
    Invokes a scriptblock to set WMI permissions, interactively or as a PSJob, with a huge HT to Graeme Bray & Steve Lee
    Created because this stuff can't be set via GPO, and the original scripts wouldn't run as remote scriptblocks/jobs, 
    making it very time consuming to run them individually/serially
    Now performs domain account lookup at the onset (to avoid double-hop problems, plus it's more efficient), and the
    scriptblock returns a comprehensive output object
    Accepts pipeline input
    Optionally tests connectivity to remote computers before invoking scriptblock
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Set-PKWMIPermissions.ps1
    Created : 2019-10-11
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-10-11 - Created script based on Graeme Bray's script from 2018-04-09 (itself modified from Steve Lee's)
        
.LINK
    https://gallery.technet.microsoft.com/Set-WMI-Namespace-Security-5081ad6d
    https://blogs.msdn.microsoft.com/wmi/2009/07/20/scripting-wmi-namespace-security-part-1-of-3/
    https://blogs.msdn.microsoft.com/wmi/2009/07/23/scripting-wmi-namespace-security-part-2-of-3/
    https://blogs.msdn.microsoft.com/wmi/2009/07/27/scripting-wmi-namespace-security-part-3-of-3/

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER Account
    Account/identity to add or remove, such as 'jbloggs' or 'DOMAIN\IT_Group' or '.\AccountName' - domain users or groups MUST use syntax 'DOMAIN\jbloggs', not 'jbloggs@domain.com' (UPN)

.PARAMETER AccountType
    Account/identity type: User or Group

.PARAMETER AccountContext
    Account/identity context: Domain or Machine (if not specified, script will try to guess based on Account syntax)

.PARAMETER Namespace
    WMI namespace (default is "@('root/cimv2','root/default','root/WMI'))"

.PARAMETER Operation
    Operation: Add or Delete permissions for account/identity

.PARAMETER Permissions
    One or more permission(s): Enable, MethodExecute, FullWrite, PartialWrite, ProviderWrite, RemoteAccess, ReadSecurity, WriteSecurity (default is *All*)

.PARAMETER Deny
    Set AccessControlEntry to 'Deny'

.PARAMETER NoInheritance
    Block inheritance of permissions settings and apply to top level of namespace only (default is inheritance)

.PARAMETER Credential
    Valid credentials on target (default is passthrough)

.PARAMETER Authentication
    WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'WMI')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> Set-PKWMIPermissions -ComputerName devbox.domain.local -Credential $Credential -Account DOMAIN\wmi_readers -AccountType Group -AccountContext Domain -Operation Add -Permissions *All* -Verbose -OutVariable really
    
        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        ComputerName   {devbox.domain.local}  
        Credential     System.Management.Automation.PSCredential
        Account        DOMAIN\wmi_readers                 
        AccountType    Group                                    
        AccountContext Domain                                   
        Operation      Add                                      
        Permissions    {*All*}                                  
        Verbose        True                                     
        OutVariable    really                                   
        Namespace      {root/cimv2, root/default, root/WMI}     
        Deny           False                                    
        AsJob          False                                    
        JobPrefix      WMIPerms                                 
        ConnectionTest WinRM                                    
        Quiet          False                                    
        ScriptName     Set-PKWMIPermissions                     
        ScriptVersion  1.0.0                                    
        PipelineInput  False                                    

        WARNING: This script presumes that the target computers are members of domain 'DOMAIN' or are in a trust relationship
        VERBOSE: [Prerequisites] Successfully found group 'wmi_readers' in domain 'DOMAIN'

        BEGIN  : Invoke scriptblock

        [devbox.domain.local] Test WinRM connection
        [devbox.domain.local] Invoke command


        Computername   : DEVBOX
        IsSuccess      : True
        Account        : wmi_readers
        AccountType    : Group
        AccountContext : Domain
        Domain         : DOMAIN
        Namespace      : root/cimv2
        Operation      : Add
        Permissions    : {Enable, MethodExecute, FullWrite, PartialWrite...}
        Deny           : False
        Messages       : {GetSecurityDescriptor succeeded, Created ACE and ACL, System.Management.ManagementBaseObject, SetSecurityDescriptor 
                         operation succeeded}

        Computername   : DEVBOX
        IsSuccess      : True
        Account        : wmi_readers
        AccountType    : Group
        AccountContext : Domain
        Domain         : DOMAIN
        Namespace      : root/default
        Operation      : Add
        Permissions    : {Enable, MethodExecute, FullWrite, PartialWrite...}
        Deny           : False
        Messages       : {GetSecurityDescriptor succeeded, Created ACE and ACL, System.Management.ManagementBaseObject, SetSecurityDescriptor 
                         operation succeeded}

        Computername   : DEVBOX
        IsSuccess      : True
        Account        : wmi_readers
        AccountType    : Group
        AccountContext : Domain
        Domain         : DOMAIN
        Namespace      : root/WMI
        Operation      : Add
        Permissions    : {Enable, MethodExecute, FullWrite, PartialWrite...}
        Deny           : False
        Messages       : {GetSecurityDescriptor succeeded, Created ACE and ACL, System.Management.ManagementBaseObject, SetSecurityDescriptor 
                         operation succeeded}

        END    : Invoke scriptblock



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
    [Alias("Computer","HostName","FQDN")]
    [object[]] $ComputerName,

    [parameter(
        Mandatory = $True,
        HelpMessage = "Account/identity to add or remove, such as 'jbloggs' or 'DOMAIN\IT_Group' or '.\AccountName' - domain users or groups MUST use syntax 'DOMAIN\jbloggs', not 'jbloggs@domain.com' (UPN)"
    )]
    [Alias("Identity")]
    [string] $Account,

    [Parameter(
        Mandatory = $True,
        HelpMessage = "Account/identity type: User or Group"
    )]
    [ValidateSet("User","Group")]
    [String]$AccountType,

    [Parameter(
        HelpMessage = "Account/identity context: Domain or Machine (if not specified, script will try to guess based on Account syntax)"
    )]
    [ValidateSet("Domain","Machine")]
    [String]$AccountContext,

    [Parameter(
        HelpMessage  = "WMI namespace (default is '@('root/cimv2','root/default','root/WMI'))'"
    )]
    [ValidateNotNullOrEmpty()]
    [String[]] $Namespace = @("root/cimv2","root/default","root/WMI"),
    
    [Parameter(
        Mandatory = $True,
        HelpMessage = "Operation: Add or Delete permissions for account/identity"
    )]
    [ValidateSet("Add","Delete")]
    [String] $Operation,
 
    [Parameter(
        HelpMessage = "One or more permission(s): Enable, MethodExecute, FullWrite, PartialWrite, ProviderWrite, RemoteAccess, ReadSecurity, WriteSecurity (default is *All*)"
    )]
    [ValidateSet("*All*","Enable","MethodExecute","FullWrite","PartialWrite","ProviderWrite","RemoteAccess","ReadSecurity","WriteSecurity")]
    [string[]] $Permissions,
 
    [Parameter(
        HelpMessage = "Set AccessControlEntry to 'Deny'"
    )]
    [Switch]$Deny,

    [Parameter(
        HelpMessage = "Block inheritance of permissions settings and apply to top level of namespace only (default is inheritance)"
    )]
    [Switch]$NoInheritance,

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
        HelpMessage = "Prefix for job name (default is 'WMI')"
    )]
    [String] $JobPrefix = "WMI",

    [Parameter(
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM)"
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

    # Function to write an error message (as a string with no stacktrace info)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    # Function to get/normalize account name and context
    Function Get-AccountInfo($Account){
        
        If (($Account -notmatch "\\|@") -and ($Account -match ".\|BUILTIN")) {$Type = "Machine"}
        If ($Account -match "\\|@") {$Type = "Domain"}
        Else {$Type = "Machine"}

        If ($Account.Contains('\')) {
            $DomainAccount = $Account.Split('\')
            $Domain = $DomainAccount[0]
            $Accountname = $DomainAccount[1]
        }
        Elseif ($Account.Contains('@')) {
            $DomainAccount = $Account.Split('@')
            $Domain = $DomainAccount[1].Split('.')[0]
            $accountname = $DomainAccount[0]
        }
        Else {
            $AccountName = $Account
            $Domain = $Null
        } 
        [PSCustomObject]@{
            Name    = $AccountName            
            Domain  = $Domain
            Type    = $Type
        }
    }

    # Function to find the domain account - if it's a local account we will do it in the scriptblock
    Function Get-DomainAccount($AccountName,$Domain,$Class){
        $Param_GetAcct = @{}
        $Param_GetAcct = @{
            Class = $Class
            Filter="Domain='$Domain' and Name='$AccountName'"
        } 
        Try {
            If ($Win32Account = Get-WmiObject @Param_GetAcct -EA SilentlyContinue) {
                [PSCustomObject]@{
                    IsFound = $True
                    Name = $AccountName
                    Domain = $Domain
                    Win32Account = $Win32Account
                }
            }
            Else {
                [PSCustomObject]@{
                    IsFound = $False
                    Name = $AccountName
                    Domain = $Domain
                    Win32Account = $Null
                }
            }
        }
        Catch {}
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

    #region Parameter checks 

    # Set all permissions if we chose All
    If ($Permissions -eq "*All*") {
        $Permissions = "Enable","MethodExecute","FullWrite","PartialWrite","ProviderWrite","RemoteAccess","ReadSecurity","WriteSecurity"
    }

    # Make sure we either have, or don't have, permissions, based on operation
    Switch ($Operation) {
        Add {
            If (-not $Permissions) {
                $Msg = "-Permissions must be specified for an Add operation"
                Write-Error $Msg
                Break
            }
        }
        Delete {
            If ($Permissions) {
                $Msg = "-Permissions cannot be specified for a Delete operation"
                Write-Error $Msg
                Break
            }
        }
    }

    # Get the domain\accountname and detected type/context
    If (-not ($AccountProperties = Get-AccountInfo -Account $Account)) {
        $Msg = "Error parsing account name $Account"
        Write-Error $Msg
        Break
    }

    # Get the WMI class
    Switch ($AccountType){
        Group {
            $Class = "Win32_Group"
            $ObjType = "group"
        }
        User  {
            $Class = "Win32_Account"
            $ObjType = "user"
        }
    } 

    Switch ($Operation) {
        Add {$ConfirmStr = "to"}
        Delete {$ConfirmStr = "on"}
    }

    # Make sure we have the correct context
    [switch]$IsDomain = $True
    Switch ($AccountContext) {
        Domain  {
            If ($AccountProperties.Type -ne "Domain") {
                $IsDomain = $False
                $Msg = "'$Account' does not appear to be a domain account; please enter as 'DOMAIN\account' or 'account@domain.local' (or select 'Machine' for AccountContext)"
                $Host.UI.WriteErrorLine("ERROR: [Prerequisites] $Msg")
                Break
            }
            Else {
                $Domain = $AccountProperties.Domain
                $AccountName = $AccountProperties.Name
            }    
        }
        Machine {
            If ($AccountProperties.Type -ne "Machine") {
                $Msg = "'$Account' does not appear to be a local machine account; please enter as 'account', '.\account', or 'BUILTIN\account' (or select 'Machine' for AccountContext)"
                $Host.UI.WriteErrorLine("ERROR: [Prerequisites] $Msg")
                Break
            }
            Else {
                $IsDomain = $False
                $Domain = $Null
                $AccountName = $AccountProperties.Name
            }
        }
        Default {
            $Domain = $AccountProperties.Domain
            $AccountName = $AccountProperties.Name
            $AccountContext = $AccountProperties.Type
            If ($AccountContext -eq "Domain") {$IsDomain = $True}
            $Msg = "Setting -AccountContext to '$AccountContext' based on account name syntax"
            Write-Verbose "[Prerequisites] $Msg"
        }
    } # end switch for account context

    # Look up the account here to avoid double-hop issues on the remote computer (plus this is more efficient)
    If ($IsDomain.IsPresent) {
        $Msg = "This script presumes that the target computers are members of domain '$Domain' or are in a trust relationship"
        Write-Warning $Msg

        If ($DomainAccount = Get-DomainAccount -AccountName $AccountName -Domain $Domain -Class $Class) {
            $Msg = "Successfully found $ObjType '$Accountname' in domain '$Domain'"
            Write-Verbose "[Prerequisites] $Msg"
        }
        Else {
            $Msg = "WMI lookup failed for $ObjType '$AccountName' in domain '$Domain'"
            $Host.UI.WriteErrorLine("ERROR: [Prerequisites] $Msg")
            Break
        }
    }

    #endregion Parameter checks    
        
    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        Param($Param_SB)
        $AccountName      = $Param_SB.AccountName
        $AccountType      = $Param_SB.AccountType
        $Class            = $Param_SB.Class
        $AccountContext   = $Param_SB.AccountContext
        $Domain           = $Param_SB.Domain
        $Namespace        = $Param_SB.NameSpace
        $Operation        = $Param_SB.Operation
        $Permissions      = $Param_SB.Permissions
        $Deny             = $Param_SB.Deny
        $NoInheritance    = $Param_SB.NoInheritance

        # Set flag
        [switch]$Continue = $False

        # Look up the local account, or accept the domain account
        Switch ($AccountContext) {
            Machine {
                $Domain = $Env:ComputerName
                $Param_GetLocalAcct = @{}
                $Param_GetLocalAcct = @{
                    Class = $Class
                    Filter="Domain='$Domain' and Name='$AccountName'"
                } 
                Try {
                    If ($Win32Account = Get-WmiObject @Param_GetLocalAcct -EA SilentlyContinue) {
                        $Continue = $True
                        $Msg = "Verified local account"
                        $Messages += $Msg

                        # Reset flag
                        $Continue = $True
                    }
                    Else {
                        $Msg = "Failed to find local account"
                        $Messages += $Msg
                    }
                }
                Catch {
                    $Msg = "Failed to find local account"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    $Messages += $Msg
                }
            }
            Domain  {
                $Win32Account = $Param_SB.Win32Account
                
                # Reset flag
                $Continue = $True
            }
        }

        # If we have a valid Win32Account
        If ($Continue.IsPresent) {
            
            # Collect all output objects (one per namespace)
            $AllOutput = @()

            # Create object for output (use this method + select for order, in case of downlevel PS version
            $OutputTemplate = New-Object PSObject -Property @{
                Computername   = $Env:ComputerName
                IsSuccess      = "Error"
                Account        = $AccountName
                AccountType    = $AccountType
                AccountContext = $AccountContext
                Domain         = $Domain
                Namespace      = $Null
                AccountSID     = $Win32Account.SID
                Inheritance    = (-not $NoInheritance.IsPresent)
                Operation      = $Operation
                Permissions    = $Permissions
                Deny           = $Deny
                Messages       = $Null
            }
            $Select = 'Computername','IsSuccess','Account','AccountType','AccountContext','Domain','Namespace','AccountSID','Inheritance','Operation','Permissions','Deny','Messages'

            # Inner function
            Function Get-AccessMaskFromPermission($permissions) {
                $WBEM_ENABLE            = 1
                $WBEM_METHOD_EXECUTE    = 2
                $WBEM_FULL_WRITE_REP    = 4
                $WBEM_PARTIAL_WRITE_REP = 8
                $WBEM_WRITE_PROVIDER    = 0x10
                $WBEM_REMOTE_ACCESS     = 0x20
                $WBEM_RIGHT_SUBSCRIBE   = 0x40
                $WBEM_RIGHT_PUBLISH     = 0x80
                $READ_CONTROL           = 0x20000
                $WRITE_DAC              = 0x40000
        
                $WBEM_RIGHTS_FLAGS = $WBEM_ENABLE,$WBEM_METHOD_EXECUTE,$WBEM_FULL_WRITE_REP,
                $WBEM_PARTIAL_WRITE_REP,$WBEM_WRITE_PROVIDER,$WBEM_REMOTE_ACCESS,$READ_CONTROL,$WRITE_DAC

                $WBEM_RIGHTS_STRINGS = "Enable","MethodExecute","FullWrite","PartialWrite","ProviderWrite",
                "RemoteAccess","ReadSecurity","WriteSecurity"
  
                $permissionTable = @{}
  
                for ($i = 0; $i -lt $WBEM_RIGHTS_FLAGS.Length; $i++) {
                    $permissionTable.Add($WBEM_RIGHTS_STRINGS[$i].ToLower(), $WBEM_RIGHTS_FLAGS[$i])
                }
                $accessMask = 0
                foreach ($permission in $permissions) {
                    if (-not $permissionTable.ContainsKey($permission.ToLower())) {
                        throw "Unknown permission: $permission`nValid permissions: $($permissionTable.Keys)"
                    }
                    $accessMask += $permissionTable[$permission.ToLower()]
                }
                $accessMask
            } #end function Get-AccessMaskFromPermissions

            Foreach ($Name in $NameSpace) {
            
                $OutputObject = $OutputTemplate.PSObject.Copy()
                $OutputObject.Namespace = $Name
                $Messages = @()

                # Reset flag
                [switch]$Continue = $False

                Try {

                    # Splat for Invoke-WMIMethod, used in multiple places
                    $Param_Get = @{}
                    $Param_Get = @{
                        Namespace = $Name
                        Path      = "__systemsecurity=@"
                    } 
                    $GetSecDesc = Invoke-WmiMethod @Param_Get -Name GetSecurityDescriptor -EA SilentlyContinue

                    # If we found the namespace
                    If ($GetSecDesc.ReturnValue -eq 0) {    
                
                        # reset flag
                        $Continue = $True

                        $Msg = "GetSecurityDescriptor succeeded"
                        $Messages += $Msg

                        $ACL = $GetSecDesc.Descriptor

                        $OBJECT_INHERIT_ACE_FLAG = 0x1
                        $CONTAINER_INHERIT_ACE_FLAG = 0x2

                    }
                    Else {
                        $Msg = "GetSecurityDescriptor failed - returnvalue $($GetSecDesc.ReturnValue)"
                        $Messages += $Msg                        
                    }
                }
                Catch {
                    $Msg = "GetSecurityDescriptor failed"
                    $Messages += $Msg    
                }

                # Execute changes
                If ($Continue.IsPresent) {
            
                    # Reset flag
                    $Continue = $False
            
                    Switch ($Operation) {
                        Add {
                            $AccessMask = Get-AccessMaskFromPermission($Permissions)
                    
                            $ACE = (New-Object System.Management.ManagementClass("win32_Ace")).CreateInstance()
                            $ACE.AccessMask = $AccessMask
                            
                            #$ACE.AceFlags = $CONTAINER_INHERIT_ACE_FLAG
                            If ($NoInheritance.IsPresent) {
                                $ACE.AceFlags = 0
                            }
                            Else {
                                $ACE.AceFlags = $OBJECT_INHERIT_ACE_FLAG + $CONTAINER_INHERIT_ACE_FLAG
                            } 
                            
                            $Trustee = (New-Object System.Management.ManagementClass("win32_Trustee")).CreateInstance()
                            $Trustee.SidString = $Win32Account.Sid
                            $ACE.Trustee = $Trustee
            
                            $ACCESS_ALLOWED_ACE_TYPE = 0x0
                            $ACCESS_DENIED_ACE_TYPE = 0x1
  
                            If ($Deny.IsPresent) {
                                $ACE.AceType = $ACCESS_DENIED_ACE_TYPE
                            } 
                            Else {
                                $ACE.AceType = $ACCESS_ALLOWED_ACE_TYPE
                            }
  
                            $ACL.DACL += $ACE.psobject.immediateBaseObject

                            # If we have everything we need to proceed
                            If ($ACL.DACL) {
                                $Msg = "Created ACE and ACL"
                                $Messages += $Msg
                                $Continue = $True
                            }
                            Else {
                                $Msg = "Failed to create ACE and ACL"
                                $Messages += $Msg
                            }                
                        } #end if operation is Add
        
                        Delete {
                            [System.Management.ManagementBaseObject[]]$NewDACL = @()
                            Foreach ($ACE in $ACL.DACL) {
                                If ($ACE.Trustee.SidString -ne $Win32account.Sid) {
                                    $NewDACL += $ACE.PSObject.immediateBaseObject
                                }
                            }
                            $ACL.DACL = $newDACL.PSObject.immediateBaseObject
                    
                            # If we have everything we need to proceed
                            If ($ACL.DACL) {
                                $Msg = "Created ACE and ACL"
                                $Messages += $Msg
                                $Continue = $True
                            }
                            Else {
                                $Msg = "Failed to created ACE and ACL"
                                $Messages += $Msg
                            }
                                
                        }
                    } #end switch to create ACE, ACL based on operation

                    If ($Continue.IsPresent) {
                            
                        Try {

                            $Param_Set = @{}
                            $Param_Set = @{
                                Name         = "SetSecurityDescriptor"
                                ArgumentList = $ACL.psobject.immediateBaseObject
                            } 
                            $Param_Set += $Param_Get

                            $SetSecDesc = Invoke-WmiMethod @Param_Set -EA Stop

                            $Messages += $SetSecDesc

                            If ($SetSecDesc.ReturnValue -eq 0) {
                                $OutputObject.IsSuccess = $True
                                $Msg = "SetSecurityDescriptor operation succeeded"
                                $Messages += $Msg
                            }
                            Else {
                                $OutputObject.IsSuccess = $False
                                $Msg = "SetSecurityDescriptor operation failed; returnvalue $($SetSecDesc.ReturnValue)"
                                $Messages += $Msg
                            }
                        }
                        Catch {
                            $Msg = "SetSecurityDescriptor operation failed"
                            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                            $Messages += $Msg
                        }
                        
                    } #end if continue to create ACE/ACL
                
                } #end if continue to execute changes
                
                # Update output object & add it to the array
                $OutputObject.Messages = $Messages
                $AllOutput += $OutputObject

            }# end foreach namespace

        } # end if we have Win32Account
        Else {
            $OutputObject = $OutputTemplate.PSObject.Copy() 
            $Msg = "Failed to find local account"
            $OutputObject.IsSuccess = $Error
            $OutputObject.Messages = $Msg
            $AllOutput += $OutputObject
        }

        # Return the array, with properties in the desired order
        Write-Output $AllOutput | Select-Object $Select
       
    } #end scriptblock

    #endregion Scriptblock for Invoke-Command

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

    # Hashtable for argumentlist parameter splat
    $Param_SB = @{}
    $Param_SB = @{
        AccountName    = $AccountName
        AccountType    = $AccountType
        Class          = $Class
        AccountContext = $AccountContext
        Domain         = $Domain
        Namespace      = $NameSpace
        Operation      = $Operation
        Permissions    = $Permissions
        Deny           = $Deny
        NoInheritance  = $NoInheritance
    }
    If ($AccountContext -eq "Domain") {
        $Param_SB.Add("Win32Account",$DomainAccount.Win32Account)
    }
    
    # Parameters for Invoke-Command
    If ($Deny.IsPresent) {
        $ConfirmMsg = "`n`n`t$Operation 'deny' permissions $($Permissions -join(", "))`t`n`tin namespace(s) $($Namespace -join(", "))`n`t$ConfirmStr $($AccountContext.ToLower()) account $Account`n`n"
    }
    Else {
        $ConfirmMsg = "`n`n`t$Operation permissions $($Permissions -join(", "))`n`tin namespace(s) $($Namespace -join(", "))`n`t$ConfirmStr $($AccountContext.ToLower()) account $Account`n`n"
    }
    $ArgumentList = $Param_SB
    $Param_IC = @{}
    $Param_IC = @{
        ComputerName   = $Null
        Authentication = $Authentication
        ArgumentList   = $ArgumentList
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

                    If ($PSCmdlet.ShouldProcess($Computer,"`n`n`t$Msg`n`n")) {
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
                    $Msg = "Test WinRM connection using authentication protocol '$Authentication'"
                    "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
                    $Param_WP.CurrentOperation = $Msg
                    Write-Progress @Param_WP

                    If ($PSCmdlet.ShouldProcess($Computer,"`n`n`t$Msg`n`n")) {
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
                        Invoke-Command @Param_IC | Select -Property * -ExcludeProperty PSComputerName,RunspaceID
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

} # end Set-PKWMIPermissions


