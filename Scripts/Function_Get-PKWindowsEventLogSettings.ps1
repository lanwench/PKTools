#requires -version 3
Function Get-PKWindowsEventLogSettings {
<#
.SYNOPSIS
    Returns information on Windows event logs on one or more computers

.DESCRIPTION
    Uses Get-WinEvent to return information on Windows event logs on one or more computers
    Details include log name, type, logmode, filesize, path, record count, and read/write dates
    Runs connectivity tests on remote computers (Ping, RPC connection on TCP port 135, or none)
    Defaults to current users' credentials; ignores credentials on local computer
    Accepts pipeline input
    Returns a PSObject

.NOTES        
    Name    : Function_Get-PKWindowsEventLogSettings.ps1
    Created : 2022-02-22
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2022-02-22 - Created script
        
.PARAMETER ComputerName
    One or more Windows computer names or objects (default is local computer)

.PARAMETER LogName
    One or more event log names (default is Application, Security, and System; you may select * for all event logs if you are a glutton)

.PARAMETER Credential
    Valid credentials on computer (ignored if local; note that authentication failure will still display a generic RPC error)

.PARAMETER ConnectionTest
    Connectivity test to run before attempting Get-WinEvent: Ping, RPC test on TCP 135, or None (default is Ping)
    
.EXAMPLE
    PS C:\> Get-PKWindowsEventLogSettings -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        Verbose        True                                     
        ComputerName   LAPTOP15                          
        LogName        {Application, Security, System}          
        Credential     System.Management.Automation.PSCredential
        ConnectionTest Ping                                     
        ScriptName     Get-PKWindowsEventLogSettings            
        ScriptVersion  1.0.0                                    
        PipelineInput  False                                    


        WARNING: Running Get-WinEvent on remote computers requires valid credentials and an enabled inbound firewall rule for 'Remote Event Log Management.'
        You can enable this GPO or via PowerShell (modify profile as needed)
	        PS C:\> Get-NetFirewallRule -Name 'RemoteEventLogSvc-In-TCP' | Set-NetFirewallRule -Enabled True -Profile Domain -Direction Inbound -PassThru
        Note also that authentication failures will still display a generic RPC connection failure.

        VERBOSE: [BEGIN: Get-PKWindowsEventLogSettings] Get Windows Event Log settings
        VERBOSE: [LAPTOP15] Getting Application event log
        VERBOSE: [LAPTOP15] Getting Security event log
        VERBOSE: [LAPTOP15] Getting System event log


        ComputerName       : LAPTOP15
        LogName            : Application
        LogType            : Administrative
        IsClassicLog       : True
        LogMode            : Circular
        FileSize           : 33558528
        IsEnabled          : True
        IsLogFull          : False
        MaximumSizeInBytes : 33554432
        RecordCount        : 69837
        LastAccessTime     : 2022-02-23 4:12:08 PM
        LastWriteTime      : 2022-02-23 4:11:52 PM
        LogFilePath        : %SystemRoot%\System32\Winevt\Logs\Application.evtx
        Messages           : 

        ComputerName       : LAPTOP15
        LogName            : Security
        LogType            : Administrative
        IsClassicLog       : True
        LogMode            : Circular
        FileSize           : 201330688
        IsEnabled          : True
        IsLogFull          : False
        MaximumSizeInBytes : 201326592
        RecordCount        : 218616
        LastAccessTime     : 2022-02-23 4:12:25 PM
        LastWriteTime      : 2022-02-23 4:11:59 PM
        LogFilePath        : %SystemRoot%\System32\Winevt\Logs\Security.evtx
        Messages           : 

        ComputerName       : LAPTOP15
        LogName            : System
        LogType            : Administrative
        IsClassicLog       : True
        LogMode            : Circular
        FileSize           : 33558528
        IsEnabled          : True
        IsLogFull          : False
        MaximumSizeInBytes : 33554432
        RecordCount        : 64476
        LastAccessTime     : 2022-02-23 4:11:04 PM
        LastWriteTime      : 2022-02-23 4:11:04 PM
        LogFilePath        : %SystemRoot%\System32\Winevt\Logs\System.evtx
        Messages           : 

        VERBOSE: [END: Get-PKWindowsEventLogSettings] Get Windows Event Log settings

.EXAMPLE
    PS C:\> "sqlserver.domain.local","dev-web-1.domain.local","192.168.25.40" | Get-PKWindowsEventLogSettings -LogName System -ConnectionTest RPC -Credential (Get-Credential)

        
        ComputerName       : sqlserver.domain.local
        LogName            : System
        LogType            : Administrative
        IsClassicLog       : True
        LogMode            : Circular
        FileSize           : 107024384
        IsEnabled          : True
        IsLogFull          : False
        MaximumSizeInBytes : 204800000
        RecordCount        : 322864
        LastAccessTime     : 2022-02-23 4:06:17 PM
        LastWriteTime      : 2022-02-23 4:06:17 PM
        LogFilePath        : %SystemRoot%\System32\Winevt\Logs\System.evtx
        Messages           : 

        ComputerName       : dev-web-1.domain.local
        LogName            : System
        LogType            : Error
        IsClassicLog       : Error
        LogMode            : Error
        FileSize           : Error
        IsEnabled          : Error
        IsLogFull          : Error
        MaximumSizeInBytes : Error
        RecordCount        : Error
        LastAccessTime     : Error
        LastWriteTime      : Error
        LogFilePath        : Error
        Messages           : Operation failed (The RPC server is unavailable)

        ComputerName       : 192.168.25.40
        LogType            : System
        IsClassicLog       : Error
        LogMode            : Error
        FileSize           : Error
        IsEnabled          : Error
        IsLogFull          : Error
        MaximumSizeInBytes : Error
        RecordCount        : Error
        LastAccessTime     : Error
        LastWriteTime      : Error
        LogFilePath        : Error 
        Messages           : RPC connection test failed or timed out after 1000 milliseconds

#>
[CmdletBinding()]
Param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "One or more computer names"
    )]
    [Alias("Identity","Name")]
    [ValidateNotNullOrEmpty()]
    [object[]]$ComputerName,

    [Parameter(
        HelpMessage = "One or more event log names (default is Application, Security, and System; you can select * for all)"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$LogName = @("Application","Security","System"),

    [Parameter(
        HelpMessage = "Valid credentials on computer (ignored if local; note that authentication failure will still display a generic RPC error)"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential]$Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(
        HelpMessage = "Connectivity test to run before attempting Get-WinEvent: Ping, RPC test on TCP 135, or None (default is Ping)"
    )]
    [ValidateSet("Ping","RPC","None")]
    [string]$ConnectionTest = "Ping"
)
Begin{
        
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $ScriptName = $MyInvocation.MyCommand.Name
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = $Env:ComputerName 
        # Doing this instead of setting a default in the parameter, because that will always display in the verbose CurrentParams even if we are using pipeline input!
    }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Msg = "Running Get-WinEvent on remote computers requires valid credentials and an enabled inbound firewall rule for 'Remote Event Log Management.'`nYou can enable this GPO or via PowerShell (modify profile as needed)`n`tPS C:\> Get-NetFirewallRule -Name 'RemoteEventLogSvc-In-TCP' | Set-NetFirewallRule -Enabled True -Profile Domain -Direction Inbound -PassThru`nNote also that authentication failures will still display a generic RPC connection failure."
    Write-Warning $Msg

    #region Select

    $Select = @{N="ComputerName";E={$Computer}},
    "LogName",
    "LogType",
    "IsClassicLog",
    "LogMode",
    "FileSize",
    "IsEnabled",
    "IsLogFull",
    "MaximumSizeInBytes",
    "RecordCount",
    "LastAccessTime",
    "LastWriteTime",
    "LogFilePath",
    @{N="Messages";E={$ResultMessage}}

    #endregion Select

    $Activity = "Get Windows Event Log settings"
    $Msg = "[BEGIN: $Scriptname] $Activity" 
    Write-Verbose $Msg
}

Process {
    
    $Total = $ComputerName.Count
    $Current = 0        
    Foreach ($Computer in $ComputerName) {
            
        $Current ++
        $Computer = $Computer.Trim()
        $Results = @()
        $ResultMessage = @()

        $Param_GWE = @{}
        $Param_GWE = @{
            ComputerName = $Computer
            ListLog      = "*"
            Verbose      = $False
            ErrorAction  = "Stop"
        }
        
        If ($Computer -match "Localhost|127.0.0.1|^$Env:ComputerName$") {
            [switch]$Continue = $True
        }
        Else {
            $Param_GWE.Add("Credential",$Credential)
            Switch ($ConnectionTest) {
                Ping {
                    $Msg = "Pinging computer"
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                    [switch]$Continue = $False
                    If ($Null = Test-Connection $Computer -Quiet -Count 1 -ErrorAction SilentlyContinue) {
                        $Continue = $True
                        $Msg = "Connection succeeded"
                        Write-Verbose "[$Computer] $Msg"
                    }
                    Else {
                        $Msg = "Ping failed"
                        Write-Warning "[$Computer] $Msg"
                        $ResultMessage += $Msg
                        $Output = "" | Select-Object $Select
                        $Output.LogName = $($LogName -join(","))
                        $Output.PSObject.Properties | Where-Object {(-not $_.Value) -or ($_.Value -eq 0)} | Foreach-Object {$_.Value = "Error"}
                        $Results += $Output
                    }
                }
                RPC {
                    $Msg = "Testing connection on TCP port 135 for RPC"
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                    [switch]$Continue = $False
                    # Using this method because it's waaaay faster than Test-NetConnection
                    If (($Status = (New-Object System.Net.Sockets.TcpClient).ConnectAsync($Computer,135).Wait(1000))-eq $True) {
                        $Continue = $True
                        $Msg = "Connection succeeded"
                        Write-Verbose "[$Computer] $Msg"
                    }
                    Else {
                        $Msg = "RPC connection test failed or timed out after 1000 milliseconds"
                        Write-Warning "[$Computer] $Msg"
                        $ResultMessage += $Msg
                        $Output = "" | Select-Object $Select
                        $Output.LogName = $($LogName -join(","))
                        $Output.PSObject.Properties | Where-Object {(-not $_.Value) -or ($_.Value -eq 0)} | Foreach-Object {$_.Value = "Error"}
                        $Results += $Output
                    }
                }
                None {
                    [switch]$Continue = $True
                }
            }
        }
        If ($Continue.IsPresent) {
            
            $Param_GWE.ComputerName = $Computer

            If ($LogName -eq "*") {
                        
                $Msg = "Getting all event logs"
                Write-Verbose "[$Computer] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                        
                Try {
                    If (($Output = Get-Winevent @Param_GWE).Count -gt 0) {
                        $ResultMessage = $Null
                        $Results += $Output
                    }
                    Else {
                        $ResultMessage = "No matching event logs found"
                        If ($ErrorDetails = $_.Exception.Message) {$ResultMessage += " ($ErrorDetails)"}
                        $Output = ("" | Select-Object $Select)
                        $Output.LogName = $($LogName)
                        $Output.PSObject.Properties | Where-Object {(-not $_.Value) -or ($_.Value -eq 0)} | Foreach-Object {$_.Value = "Error"}
                        $Results += $Output
                    }
                }
                Catch {
                    $ResultMessage = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$ResultMessage += " ($ErrorDetails)"}
                    $Output = ("" | Select-Object $Select)
                    $Output.LogName = $($LogName)
                    $Output.PSObject.Properties | Where-Object {(-not $_.Value) -or ($_.Value -eq 0)} | Foreach-Object {$_.Value = "Error"}
                    $Results += $Output
                }
            }
            Else {
                Foreach ($Log in $LogName) {
                            
                    $Msg = "Getting $Log event log"
                    Write-Verbose "[$Computer] $Msg"
                    Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete($Current/$Total*100)
                    Try {                            
                        $Param_GWE.ListLog = $Log
                        If (($Output = Get-Winevent @Param_GWE).Count -gt 0) {
                            $ResultMessage = "Operation completed successfully"
                            $Results += $Output
                        }
                        Else {
                            $ResultMessage = "Log not found"
                            If ($ErrorDetails = $_.Exception.Message) {$ResultMessage += " ($ErrorDetails)"}
                            $Output = ("" | Select-Object $Select)
                            $Output.LogName = $($Log)
                            $Output.PSObject.Properties | Where-Object {(-not $_.Value) -or ($_.Value -eq 0)} | Foreach-Object {$_.Value = "Error"}
                            $Results += $Output
                        }
                    }
                    Catch {
                        $ResultMessage = "Operation failed"
                        If ($ErrorDetails = $_.Exception.Message) {$ResultMessage += " ($ErrorDetails)"}
                        $Output = ("" | Select-Object $Select)
                        $Output.LogName = $($Log)
                        $Output.PSObject.Properties | Where-Object {(-not $_.Value) -or ($_.Value -eq 0)} | Foreach-Object {$_.Value = "Error"}
                        $Results += $Output
                    }
                }   
            }
        }
        
        Write-Output ($Results | Select-Object $Select)

    } #end foreach 
}
End {

    $Null = Write-Progress -Activity * -Completed
    $Msg = "[END: $Scriptname] $Activity" 
    Write-Verbose $Msg

}
} #end Get-PKWindowsEventLogSettings


$Null = New-Alias Get-PKWinEventLogSettings -Value Get-PKWindowsEventLogSettings -Force -Confirm:$False