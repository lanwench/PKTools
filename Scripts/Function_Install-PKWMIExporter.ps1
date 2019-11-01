#requires -version 3
Function Install-PKWMIExporter {
<# 
.SYNOPSIS
    Invokes a scriptblock to download and install WMI Exporter, with specified collectors/listening port - defaults to AD collectors

.DESCRIPTION
    Invokes a scriptblock to download and install WMI Exporter, with specified collectors/listening port - defaults to AD collectors
    Accepts pipeline input
    Optionally tests connectivity to remote computers before invoking scriptblock
    Returns a PSobject or PSJob

.NOTES        
    Name    : Install-PKWMIExporter.ps1
    Created : 2019-10-07
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-10-07 - Created script


.LINK
    https://github.com/martinlindhe/wmi_exporter

.EXAMPLE
    PS C:\> Install-PKWMIExporter -ComputerName server123.domain.local -Authentication Kerberos -Credential $Creds -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                                                                           
        ---              -----                                                                                           
        ComputerName     {server123.domain.local}                                                               
        Authentication   Kerberos                                                                                        
        Credential       System.Management.Automation.PSCredential                                                       
        Verbose          True                                                                                            
        URI              https://github.com/martinlindhe/wmi_exporter/releases/download/v0.8.3/wmi_exporter-0.8.3-386.msi
        Collectors       ad,cpu,cs,dns,logical_disk,net,os,service,system                                                
        ListenPort       9182                                                                                            
        Force            False                                                                                           
        ConnectionTest   WinRM                                                                                           
        AsJob            False                                                                                           
        JobPrefix        WMIExporter                                                                                     
        NoProgress       False                                                                                           
        Quiet            False                                                                                           
        ParameterSetName Interactive                                                                                     
        PipelineInput    False                                                                                           
        ScriptName       Install-PKWMIExporter                                                                           
        ScriptVersion    1.0.0                                                                                           

        BEGIN  : Download and install WMIExporter

        [server123.domain.local] Test WinRM connection
        [server123.domain.local] Invoke command


        ComputerName   : SERVER123
        IsSuccess      : True
        ServiceRunning : Running
        ServicePath    : "C:\Program Files (x86)\wmi_exporter\wmi_exporter.exe" --log.format logger:eventlog?name=wmi_exporter --collectors.enabled 
                         ad,cpu,cs,dns,logical_disk,net,os,service,system --telemetry.addr :9182   
        URIPath        : https://github.com/martinlindhe/wmi_exporter/releases/download/v0.8.3/wmi_exporter-0.8.3-386.msi
        LocalPath      : C:\Users\JBloggs\AppData\Local\Temp\wmi_exporter-0.8.3-386.msi
        SetupCommand   : Start-Process -FilePath "msiexec.exe" -ArgumentList /qn /i C:\Users\JBLOGGS\AppData\Local\Temp\wmi_exporter-0.8.3-386.msi 
                         /L*V C:\Users\JBLOGGS\AppData\Local\Temp\2019-10-07_18-10-20_WMIExporterLog.log 
                         ENABLED_COLLECTORS=ad,cpu,cs,dns,logical_disk,net,os,service,system LISTEN_PORT=9182 -Wait -PassThru -NoNewWindow 
                         -ErrorAction Stop
        LogFile        : C:\Users\JBloggs\AppData\Local\Temp\2019-10-07_18-10-20_WMIExporterLog.log
        Messages       : Application installed successfully - return code 0


        END    : Download and install WMIExporter

.EXAMPLE
    PS C:\> $DCs.Hostname | Install-PKWMIExporter -Credential $AdminCreds -AsJob -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                                                                           
        ---              -----                                                                                           
        Credential       System.Management.Automation.PSCredential                                                       
        AsJob            True                                                                                            
        Verbose          True                                                                                            
        ComputerName                                                                                                     
        URI              https://github.com/martinlindhe/wmi_exporter/releases/download/v0.8.3/wmi_exporter-0.8.3-386.msi
        Collectors       ad,cpu,cs,dns,logical_disk,net,os,service,system                                                
        ListenPort       9182                                                                                            
        Force            False                                                                                           
        Authentication   Negotiate                                                                                       
        ConnectionTest   WinRM                                                                                           
        JobPrefix        WMIExporter                                                                                     
        NoProgress       False                                                                                           
        Quiet            True                                                                                          
        ParameterSetName Job                                                                                             
        PipelineInput    True                                                                                            
        ScriptName       Install-PKWMIExporter                                                                           
        ScriptVersion    1.0.0                                                                                           


        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        114    WMIExporter_... RemoteJob       Completed     True            DC04.domain... ...                      
        116    WMIExporter_... RemoteJob       Completed     True            TEMPDC.domain... ...                      
        118    WMIExporter_... RemoteJob       Completed     True            DC08.domain... ...                      
        120    WMIExporter_... RemoteJob       Completed     True            DC05.domain... ...                      
        122    WMIExporter_... RemoteJob       Completed     True            DC02.domain.... ...                                        

        PS C:\> Get-Job | Receive-Job | Format-Table -Autosize

        ComputerName IsSuccess ServiceRunning ServicePath                                                                   URIPath              
        ------------ --------- -------------- -----------                                                                   -------              
        DC04             False Running        "C:\Program Files (x86)\wmi_exporter\wmi_exporter.exe" --log.format log...    https://github.com...
        TEMPDC           False Running        "C:\Program Files (x86)\wmi_exporter\wmi_exporter.exe" --log.format log...    https://github.com...
        DC08             False Error          Error                                                                         https://github.com...
        DC05             False Running        "c:\Program Files (x86)\wmi_exporter\wmi_exporter.exe" --log.format log...    https://github.com...
        DC02             False Running        "C:\Program Files (x86)\wmi_exporter\wmi_exporter.exe" --log.format log...    https://github.com...


    
#> 

[CmdletBinding(
    DefaultParameterSetName = "Interactive",
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
        HelpMessage = "Path to WMI exporter MSI file download (default is 'https://github.com/martinlindhe/wmi_exporter/releases/download/v0.8.3/wmi_exporter-0.8.3-386.msi')"
    )]
    [ValidateNotNullOrEmpty()]
    [string] $URI = "https://github.com/martinlindhe/wmi_exporter/releases/download/v0.8.3/wmi_exporter-0.8.3-386.msi",

    [Parameter(
        HelpMessage = "Collectors (default is 'ad,cpu,cs,dns,logical_disk,net,os,service,system')"
    )]
    [ValidateNotNullOrEmpty()]
    [String] $Collectors = 'ad,cpu,cs,dns,logical_disk,net,os,service,system',

    [Parameter(
        HelpMessage = "The port to bind to (default is 9182)"
    )]
    [ValidateNotNullOrEmpty()]
    [int] $ListenPort = 9182,

    [Parameter(
        HelpMessage = "Force installation even if Windows service already exists"
    )]
    [Switch] $Force,

    [Parameter(
        HelpMessage = "Valid credentials on target (default is passthrough; credential is ignored on local computer)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)"
    )]
    [ValidateSet('Kerberos','Basic','Negotiate','Default','CredSSP')]
    [string]$Authentication = "Negotiate",

    [Parameter(
        HelpMessage = "Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "WinRM",

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Run Invoke-Command scriptblock as PSJob"
    )]
    [Switch] $AsJob,

    [Parameter(
        ParameterSetName = "Job",
        HelpMessage = "Prefix for job name (default is 'WMIExporter')"
    )]
    [String] $JobPrefix = "WMIExporter",

    [Parameter(
        HelpMessage = "Don't display progress bar"
    )]
    [Switch] $NoProgress,

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
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Preferences 
    
    $ErrorActionPreference = "Stop"
    Switch ($NoProgress) {
        $True  {$ProgressPreference = "SilentlyContinue"}
        $False {$ProgressPreference = "Continue"}
    }

    #endregion Preferences
    
    #region Exit codes

    # MSIExec exit codes for lookup 
    $ExitCodes = @{
    
        [int]'0'   ='[ERROR_SUCCESS] The action completed successfully.'
        [int]'13'  ='[ERROR_INVALID_DATA] The data is invalid.'
        [int]'87'  ='[ERROR_INVALID_PARAMETER] One of the parameters was invalid.'
        [int]'120' ='[ERROR_CALL_NOT_IMPLEMENTED] This value is returned when a custom action attempts to call a function that cannot be called from custom actions. The function returns the value ERROR_CALL_NOT_IMPLEMENTED. Available beginning with Windows Installer version 3.0.'
        [int]'1259'='[ERROR_APPHELP_BLOCK] If Windows Installer determines a product may be incompatible with the current operating system, it displays a dialog box informing the user and asking whether to try to install anyway. This error code is returned if the user chooses not to try the installation.'
        [int]'1601'='[ERROR_INSTALL_SERVICE_FAILURE] The Windows Installer service could not be accessed. Contact your support personnel to verify that the Windows Installer service is properly registered.'
        [int]'1602'='[ERROR_INSTALL_USEREXIT] The user cancels installation.'
        [int]'1603'='[ERROR_INSTALL_FAILURE] A fatal error occurred during installation.'
        [int]'1604'='[ERROR_INSTALL_SUSPEND] Installation suspended, incomplete.'
        [int]'1605'='[ERROR_UNKNOWN_PRODUCT] This action is only valid for products that are currently installed.'
        [int]'1606'='[ERROR_UNKNOWN_FEATURE] The feature identifier is not registered.'
        [int]'1607'='[ERROR_UNKNOWN_COMPONENT] The component identifier is not registered.'
        [int]'1608'='[ERROR_UNKNOWN_PROPERTY] This is an unknown property.'
        [int]'1609'='[ERROR_INVALID_HANDLE_STATE] The handle is in an invalid state.'
        [int]'1610'='[ERROR_BAD_CONFIGURATION] The configuration data for this product is corrupt. Contact your support personnel.'
        [int]'1611'='[ERROR_INDEX_ABSENT] The component qualifier not present.'
        [int]'1612'='[ERROR_INSTALL_SOURCE_ABSENT] The installation source for this product is not available. Verify that the source exists and that you can access it.'
        [int]'1613'='[ERROR_INSTALL_PACKAGE_VERSION] This installation package cannot be installed by the Windows Installer service. You must install a Windows service pack that contains a newer version of the Windows Installer service.'
        [int]'1614'='[ERROR_PRODUCT_UNINSTALLED] The product is uninstalled.'
        [int]'1615'='[ERROR_BAD_QUERY_SYNTAX] The SQL query syntax is invalid or unsupported.'
        [int]'1616'='[ERROR_INVALID_FIELD] The record field does not exist.'
        [int]'1618'='[ERROR_INSTALL_ALREADY_RUNNING] Another installation is already in progress. Complete that installation before proceeding with this install. '
        [int]'1619'='[ERROR_INSTALL_PACKAGE_OPEN_FAILED] This installation package could not be opened. Verify that the package exists and is accessible, or contact the application vendor to verify that this is a valid Windows Installer package.'
        [int]'1620'='[ERROR_INSTALL_PACKAGE_INVALID] This installation package could not be opened. Contact the application vendor to verify that this is a valid Windows Installer package.'
        [int]'1621'='[ERROR_INSTALL_UI_FAILURE] There was an error starting the Windows Installer service user interface. Contact your support personnel.'
        [int]'1622'='[ERROR_INSTALL_LOG_FAILURE] There was an error opening installation log file. Verify that the specified log file location exists and is writable.'
        [int]'1623'='[ERROR_INSTALL_LANGUAGE_UNSUPPORTED] This language of this installation package is not supported by your system.'
        [int]'1624'='[ERROR_INSTALL_TRANSFORM_FAILURE] There was an error applying transforms. Verify that the specified transform paths are valid.'
        [int]'1625'='[ERROR_INSTALL_PACKAGE_REJECTED] This installation is forbidden by system policy. Contact your system administrator.'
        [int]'1626'='[ERROR_FUNCTION_NOT_CALLED] The function could not be executed.'
        [int]'1627'='[ERROR_FUNCTION_FAILED] The function failed during execution.'
        [int]'1628'='[ERROR_INVALID_TABLE] An invalid or unknown table was specified.'
        [int]'1629'='[ERROR_DATATYPE_MISMATCH] The data supplied is the wrong type.'
        [int]'1630'='[ERROR_UNSUPPORTED_TYPE] Data of this type is not supported.'
        [int]'1631'='[ERROR_CREATE_FAILED] The Windows Installer service failed to start. Contact your support personnel.'
        [int]'1632'='[ERROR_INSTALL_TEMP_UNWRITABLE] The Temp folder is either full or inaccessible. Verify that the Temp folder exists and that you can write to it.'
        [int]'1633'='[ERROR_INSTALL_PLATFORM_UNSUPPORTED] This installation package is not supported on this platform. Contact your application vendor.'
        [int]'1634'='[ERROR_INSTALL_NOTUSED] Component is not used on this machine.'
        [int]'1635'='[ERROR_PATCH_PACKAGE_OPEN_FAILED] This patch package could not be opened. Verify that the patch package exists and is accessible, or contact the application vendor to verify that this is a valid Windows Installer patch package.'
        [int]'1636'='[ERROR_PATCH_PACKAGE_INVALID] This patch package could not be opened. Contact the application vendor to verify that this is a valid Windows Installer patch package.'
        [int]'1637'='[ERROR_PATCH_PACKAGE_UNSUPPORTED] This patch package cannot be processed by the Windows Installer service. You must install a Windows service pack that contains a newer version of the Windows Installer service.'
        [int]'1638'='[ERROR_PRODUCT_VERSION] Another version of this product is already installed. Installation of this version cannot continue. To configure or remove the existing version of this product, use�Add/Remove Programs�in�Control Panel.'
        [int]'1639'='[ERROR_INVALID_COMMAND_LINE] Invalid command line argument. Consult the Windows Installer SDK for detailed command-line help.'
        [int]'1640'='[ERROR_INSTALL_REMOTE_DISALLOWED] The current user is not permitted to perform installations from a client session of a server running the Terminal Server role service.'
        [int]'1641'='[ERROR_SUCCESS_REBOOT_INITIATED] The installer has initiated a restart. This message is indicative of a success.'
        [int]'1642'='[ERROR_PATCH_TARGET_NOT_FOUND] The installer cannot install the upgrade patch because the program being upgraded may be missing or the upgrade patch updates a different version of the program. Verify that the program to be upgraded exists on your computer and that you have the correct upgrade patch.'
        [int]'1643'='[ERROR_PATCH_PACKAGE_REJECTED] The patch package is not permitted by system policy.'
        [int]'1644'='[ERROR_INSTALL_TRANSFORM_REJECTED] One or more customizations are not permitted by system policy.'
        [int]'1645'='[ERROR_INSTALL_REMOTE_PROHIBITED] Windows Installer does not permit installation from a Remote Desktop Connection.'
        [int]'1646'='[ERROR_PATCH_REMOVAL_UNSUPPORTED] The patch package is not a removable patch package. Available beginning with Windows Installer version 3.0.'
        [int]'1647'='[ERROR_UNKNOWN_PATCH] The patch is not applied to this product. Available beginning with Windows Installer version 3.0.'
        [int]'1648'='[ERROR_PATCH_NO_SEQUENCE] No valid sequence could be found for the set of patches. Available beginning with Windows Installer version 3.0.'
        [int]'1649'='[ERROR_PATCH_REMOVAL_DISALLOWED] Patch removal was disallowed by policy. Available beginning with Windows Installer version 3.0.'
        [int]'1650'='[ERROR_INVALID_PATCH_XML] The XML patch data is invalid. Available beginning with Windows Installer version 3.0.'
        [int]'1651'='[ERROR_PATCH_MANAGED_ADVERTISED_PRODUCT] Administrative user failed to apply patch for a per-user managed or a per-machine application that is in advertise state. Available beginning with Windows Installer version 3.0.'
        [int]'1652'='[ERROR_INSTALL_SERVICE_SAFEBOOT] Windows Installer is not accessible when the computer is in Safe Mode. Exit Safe Mode and try again or try using�System Restore�to return your computer to a previous state. Available beginning with  Windows Installer version 4.0.'
        [int]'1653'='[ERROR_ROLLBACK_DISABLED] Could not perform a multiple-package transaction because rollback has been disabled.�Multiple-Package Installationscannot run if rollback is disabled. Available beginning with Windows Installer version 4.5.'
        [int]'1654'='[ERROR_INSTALL_REJECTED] The app that you are trying to run is not supported on this version of Windows. A Windows Installer package, patch, or transform that has not been signed by Microsoft cannot be installed on an ARM computer.'
        [int]'3010'='[ERROR_SUCCESS_REBOOT_REQUIRED] A restart is required to complete the install. This message is indicative of a success. This does not include installs where the�ForceReboot�action is run.'
    }

    #endregion Exit codes

    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        
        Param($URI,$Collectors,$ListenPort,$Force,$ExitCodes)
        
        $Output = New-Object PSObject -Property @{
            ComputerName   = $env:COMPUTERNAME
            IsSuccess      = "Error"
            ServiceRunning = "Error"
            ServicePath    = "Error"
            ListenPort     = $ListenPort
            URIPath        = $URI
            LocalPath      = "Error"
            SetupCommand   = "Error"
            LogFile        = "Error"
            Messages       = "Error"
        }

        [switch]$Continue = $False
        
        # Invoke-WebRequest requires v3 at minimum; not using -Requires as we want output regardless
        If ($PSVersionTable.PSVersion.Major -lt 3) {
            
            $Msg = "Script requires PS3 at minimum; computer is running $($PSVersionTable.PSVersion.ToString())"

            If ($Service = Get-WmiObject win32_service -filter "Name='wmi_exporter'" -ErrorAction SilentlyContinue) {
                $Msg += " (Windows service already present)"
                $Output.ServiceRunning = $Service.State
                $Output.ServicePath = $Service.PathName
            }
            
            $Output.IsSuccess = $False
            $Output.Messages = $Msg
        }
        Else {

            # Reset flag
            $Continue = $True

            $Filename = $URI | Split-Path -Leaf
            $FilePath =  "$Env:Temp\$Filename"
            $DateStamp = get-date -Format yyyy-MM-dd_HH-mm-ss
            $Logfile = "$Env:Temp\$($DateStamp)_WMIExporterLog.log"
            $ArgumentList = "/qn /i $FilePath /L*V $LogFile ENABLED_COLLECTORS=$Collectors LISTEN_PORT=$ListenPort"
            $CommandStr = "Start-Process -FilePath ""msiexec.exe"" -ArgumentList $ArgumentList -Wait -PassThru -NoNewWindow -ErrorAction Stop"

            $Output.SetupCommand = $CommandStr
            $Output.LocalPath = $FilePath
        }
        
        If ($Continue.IsPresent) {
            
            # Reset flag
            $Continue = $False

            If ($Service = Get-WmiObject win32_service -filter "Name='wmi_exporter'" -ErrorAction SilentlyContinue) {
                If (-not $Force.IsPresent) {
                    $Msg = "Windows service already present; -Force not detected"
                    $Output.ServiceRunning = $Service.State
                    $Output.ServicePath = $Service.PathName
                    $Output.IsSuccess = $False
                    $Output.Messages = $Msg
                }
                Else {
                    $Continue = $True
                }
            }
            Else {
                $Continue = $True
            }
        }

        If ($Continue.IsPresent) {
            Try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls
                $Download = Invoke-WebRequest -Uri $URI -UseBasicParsing -OutFile $Filepath -PassThru -EA Stop -Method Get

                If ($Download.StatusCode -eq 200) {
                    
                    If ($SetupFile = Get-Item $FilePath) {
                    
                        Try {
                            $Output.LogFile = $Logfile
                            $Execute = Start-Process -FilePath "msiexec.exe" -ArgumentList $ArgumentList -Wait -PassThru -NoNewWindow -ErrorAction Stop
                        
                            If ($Execute.ExitCode -eq 0) {
                                $Service = Get-WmiObject win32_service -filter "Name='wmi_exporter'" -ErrorAction SilentlyContinue
                                $Output.IsSuccess = $True
                                $Output.ServicePath = $Service.PathName
                                $Output.ServiceRunning = $Service.State
                                $Msg = "Application installed successfully - return code $($Execute.ExitCode)"
                                $Output.Messages = $Msg
                            }
                            Else {
                                $Output.IsSuccess = $False
                                $Msg = "Application failed to install; return code $($Execute.ExitCode)"
                                $Msg += " ($($ExitCodes.Item($Execute.ExitCode))"

                                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                                $Output.Messages = $Msg
                            }
                        }
                        Catch {
                            $Msg = "Failed to install MSI"
                            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                            $Output.IsSuccess = $False
                            $Output.Messages = $Msg
                        }
                    }
                }
                Else {
                    $Msg = "Failed to download MSI; status code: $($Download.StatusCode)"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    $Output.IsSuccess = $False
                    $Output.Messages = $Msg
                }
            }
            Catch {
                $Msg = "Failed to download MSI"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                $Output.IsSuccess = $False
                $Output.Messages = $Msg
            }
        }

        Write-Output $Output | Select-Object Computername,IsSuccess,ServiceRunning,ServicePath,URIPath,LocalPath,SetupCommand,LogFile,Messages

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
        Param([Parameter(ValueFromPipeline)]$Message)
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

    # Splats for Invoke-Command (remote)
    $Param_IC_Remote = @{}
    $Param_IC_Remote = @{
        ComputerName   = $Null
        Authentication = $Authentication
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $URI,$Collectors,$ListenPort,$Force,$ExitCodes
        Credential     = $Credential
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($AsJob.IsPresent) {
        $Jobs = @()
        $Param_Job = @{}
        $Param_Job = @{
            AsJob   = $True
            JobName = $Null
        }
    }

    # Splat for Invoke-Command (local)
    $Param_IC_Local = @{}
    $Param_IC_Local = @{
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $URI,$Collectors,$ListenPort,$Force,$ExitCodes
        ErrorAction    = "Stop"
        Verbose        = $False
    }
        
    # Splat for Start-Job (local computer)
    $Param_SJ = @{}
    $Param_SJ = @{
        Authentication = "Default"
        ScriptBlock    = $ScriptBlock
        ArgumentList   = $URI,$Collectors,$ListenPort,$Force,$ExitCodes
        Name           = $Null
        ErrorAction    = "Stop"
        Verbose        = $False
    }
    If ($Credential.Username) {
        $Param_SJ.Add("Credential",$Credential)
    } 

    # Splat for Write-Progress
    $Activity = "Download and install WMIExporter"
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
    $ConfirmMsg = $Activity
    
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

                    If ($PSCmdlet.ShouldProcess($Computer,"`n`n`t$Msg`n")) {
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
            
            $Msg = "Invoke command"
            If ($AsJob.IsPresent) {$Msg += " as PSJob"}
            "[$Computer] $Msg" | Write-MessageInfo -FGColor White
                    
            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP

            If ($PSCmdlet.ShouldProcess($Computer,"`n`n`t$Activity`n`n")) {
                
                Try {
                    If ($Env:ComputerName -eq $Computer) {
                        If ($AsJob.IsPresent) {
                            $Job = $Null
                            $Param_SJ.Name = "$JobPrefix`_$Computer"
                            $Job = Start-Job @Param_SJ
                            $Jobs += $Job
                        }
                        Else {
                            Invoke-Command @Param_IC_Local | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceId
                        }
                    }
                    Else {   
                        $Param_IC_Remote.ComputerName = $Computer
                        If ($AsJob.IsPresent) {
                            $Job = $Null
                            $Param_Job.JobName = "$JobPrefix`_$Computer"
                            $Job = Invoke-Command @Param_IC_Remote @Param_Job
                            $Jobs += $Job
                        }
                        Else {
                            Invoke-Command @Param_IC_Remote | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceId
                        }   
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
            "$Msg" | Write-MessageInfo -FGColor Green -Title
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

} # end Install-PKWMIExporter



