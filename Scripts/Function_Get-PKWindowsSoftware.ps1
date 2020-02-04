#Requires -version 3
Function Get-PKWindowsSoftware {
<# 
.SYNOPSIS
    Invokes a scriptblock to return the installed software on one or more Windows computers, using registry lookups instead of WMI

.DESCRIPTION
    Invokes a scriptblock to return the installed software on one or more Windows computers, using registry lookups instead of WMI
    Accepts pipeline input
    Optionally tests connectivity to remote computers before invoking scriptblock
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsSoftware.ps1
    Created : 2020-01-22
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-01-22 - Created script based on Marc Carter's original (see link)

.LINK
    https://devblogs.microsoft.com/scripting/use-powershell-to-quickly-find-installed-software/

.PARAMETER ComputerName
    One or more computer names (default is local computer)

.PARAMETER Credential
    Valid credentials on target (default is passthrough)

.PARAMETER Authentication
    WinRM authentication protocol: Kerberos, Basic, Negotiate, Default (default is Negotiate; is ignored on local computer)

.PARAMETER AsJob
    Run Invoke-Command scriptblock as PSJob

.PARAMETER JobPrefix
    Prefix for job name (default is 'Software')

.PARAMETER ConnectionTest
    Options to test connectivity on remote computer prior to Invoke-Command - WinRM, ping, or none (default is WinRM; tests ignored on local computer)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> Get-PKWindowsSoftware -verbose
        VERBOSE: PSBoundParameters: 
	
        Key            Value                                    
        ---            -----                                    
        Verbose        True                                     
        ComputerName   RAINBOWDASH                          
        Credential     System.Management.Automation.PSCredential
        Authentication Negotiate                                
        AsJob          False                                    
        JobPrefix      Software                                 
        ConnectionTest WinRM                                    
        Quiet          False                                    
        ScriptName     Get-PKWindowsSoftware                    
        ScriptVersion  1.0.0                                    
        PipelineInput  False                                    

        BEGIN  : Invoke scriptblock

        [RAINBOWDASH] Invoke command

        ComputerName    : RAINBOWDASH
        DisplayName     : Git version 2.22.0.windows.1
        DisplayVersion  : 2.22.0.windows.1
        InstallLocation : C:\Program Files\Git\
        InstallDate     : 2019-08-14
        Publisher       : The Git Development Community
        UninstallString : "C:\Program Files\Git\unins000.exe"

        ComputerName    : RAINBOWDASH
        DisplayName     : Microsoft Office 365 ProPlus - en-us
        DisplayVersion  : 16.0.12325.20298
        InstallLocation : C:\Program Files\Microsoft Office
        InstallDate     : 
        Publisher       : Microsoft Corporation
        UninstallString : "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None 
                          productstoremove=O365ProPlusRetail.16_en-us_x-none culture=en-us version.16=16.0

        ComputerName    : RAINBOWDASH
        DisplayName     : Microsoft Visual Studio Code
        DisplayVersion  : 1.41.1
        InstallLocation : C:\Program Files\Microsoft VS Code\
        InstallDate     : 2019-12-30
        Publisher       : Microsoft Corporation
        UninstallString : "C:\Program Files\Microsoft VS Code\unins000.exe"

        ComputerName    : RAINBOWDASH
        DisplayName     : Mozilla Firefox 72.0.1 (x64 en-US)
        DisplayVersion  : 72.0.1
        InstallLocation : C:\Program Files\Mozilla Firefox
        InstallDate     : 
        Publisher       : Mozilla
        UninstallString : "C:\Program Files\Mozilla Firefox\uninstall\helper.exe"

        ComputerName    : RAINBOWDASH
        DisplayName     : TreeSize V7.1.2 (64 bit)
        DisplayVersion  : 7.1.2
        InstallLocation : C:\Program Files\JAM Software\TreeSize\
        InstallDate     : 2019-09-23
        Publisher       : JAM Software
        UninstallString : "C:\Program Files\JAM Software\TreeSize\unins000.exe"

        END    : Invoke scriptblock

.EXAMPLE
    PS C:\> Get-VM test-vm* | Get-PKWindowsSoftware -Quiet -AsJob -Credential $Credential

        Id     Name            PSJobTypeName   State         HasMoreData     Location         Command                  
        --     ----            -------------   -----         -----------     --------         -------                  
        155    Software_te-... RemoteJob       Completed     True            TEST-VM-35       ...                      
        157    Software_te-... RemoteJob       Completed     True            TEST-VM-8        ...                      
        159    Software_te-... RemoteJob       Completed     True            TEST-VM-42       ...                      
        161    Software_te-... RemoteJob       Running       True            TEST-VM-16       ...                      

        PS: C\> Get-Job soft* | Receive-job -Keep | Format-Table -AutoSize

        ComputerName   DisplayName                                                    DisplayVersion  InstallLocation                        InstallDate Publisher         
        ------------   -----------                                                    --------------  ---------------                        ----------- ---------         
        TEST-VM-35     Chef Client v14.1.12                                           14.1.12.1                                              2019-12-10  Chef Software, ...
        TEST-VM-35     Java SE Development Kit 8 Update 172 (64-bit)                  8.0.1720.11     C:\Program Files\Java\jdk1.8.0_172\    2018-04-26  Oracle Corporation
        TEST-VM-35     Microsoft Visual C++ 2008 Redistributable - x64 9.0.30729.6161 9.0.30729.6161                                         2015-02-17  Microsoft Corpo...
        TEST-VM-35     VMware Tools                                                   10.1.15.6677369 C:\Program Files\VMware\VMware Tools\  2018-04-23  VMware, Inc.      
        TEST-VM-8      Chef Client v14.1.12                                           14.1.12.1                                              2019-12-10  Chef Software, ...
        TEST-VM-8      Java SE Development Kit 8 Update 131 (64-bit)                  8.0.1310.11     C:\Program Files\Java\jdk1.8.0_131\    2017-06-27  Oracle Corporation
        TEST-VM-8      Microsoft Visual C++ 2008 Redistributable - x64 9.0.30729.6161 9.0.30729.6161                                         2015-02-17  Microsoft Corpo...
        TEST-VM-8      Microsoft Visual C++ 2012 x64 Additional Runtime - 11.0.61030  11.0.61030                                             2017-04-06  Microsoft Corpo...
        TEST-VM-8      Microsoft Visual C++ 2012 x64 Minimum Runtime - 11.0.61030     11.0.61030                                             2017-04-06  Microsoft Corpo...
        TEST-VM-8      Microsoft Visual C++ 2013 x64 Additional Runtime - 12.0.21005  12.0.21005                                             2017-05-10  Microsoft Corpo...
        TEST-VM-8      Microsoft Visual C++ 2013 x64 Minimum Runtime - 12.0.21005     12.0.21005                                             2017-05-10  Microsoft Corpo...
        TEST-VM-8      VMware Tools                                                   10.1.0.4449150  C:\Program Files\VMware\VMware Tools\  2017-03-08  VMware, Inc.      
        TEST-VM-42     Chef Client v14.1.12                                           14.1.12.1                                              2019-12-10  Chef Software, ...
        TEST-VM-42     Java SE Development Kit 8 Update 172 (64-bit)                  8.0.1720.11     C:\Program Files\Java\jdk1.8.0_172\    2018-04-26  Oracle Corporation
        TEST-VM-42     Microsoft Visual C++ 2008 Redistributable - x64 9.0.30729.6161 9.0.30729.6161                                         2015-02-17  Microsoft Corpo...
        TEST-VM-42     VMware Tools                                                   10.1.15.6677369 C:\Program Files\VMware\VMware Tools\  2018-04-23  VMware, Inc.      
        TEST-VM-16     Chef Client v14.1.12                                           14.1.12.1                                              2019-12-10  Chef Software, ...
        TEST-VM-16     Java 8 Update 162 (64-bit)                                     8.0.1620.12     C:\Program Files\Java\jre1.8.0_162\    2018-03-28  Oracle Corporation
        TEST-VM-16     Java SE Development Kit 8 Update 162 (64-bit)                  8.0.1620.12     C:\Program Files\Java\jdk1.8.0_162\    2018-03-28  Oracle Corporation
        TEST-VM-16     Microsoft Visual C++ 2008 Redistributable - x64 9.0.30729.6161 9.0.30729.6161                                         2015-02-17  Microsoft Corpo...
        TEST-VM-16     Python 3.6.4 Core Interpreter (64-bit)                         3.6.4150.0                                             2018-03-23  Python Software...
        TEST-VM-16     VMware Tools                                                   10.1.0.4449150  C:\Program Files\VMware\VMware Tools\  2017-03-08  VMware, Inc.      



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
        HelpMessage = "Prefix for job name (default is 'Software')"
    )]
    [String] $JobPrefix = "Software",

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
        
        #Define the variable to hold the location of Currently Installed Programs
        $UninstallKey=”SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall” 

        #Create an instance of the Registry Object and open the HKLM base key
        $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey(‘LocalMachine’,$Env:ComputerName) 

        #Drill down into the Uninstall key using the OpenSubKey Method
        $regkey=$reg.OpenSubKey($UninstallKey) 

        #Retrieve an array of string that contain all the subkey names
        $subkeys=$regkey.GetSubKeyNames() 

        #Open each Subkey and use GetValue Method to return the required values for each
        $Output = @()
        foreach($key in $subkeys){

            $thisKey=$UninstallKey+”\\”+$key 
            $thisSubKey=$reg.OpenSubKey($thisKey)
            #If ($thisSubKey.GetValue(“InstallDate”)) {
            #    $Date = [datetime]::parseexact($($thisSubKey.GetValue(“InstallDate”)), 'yyyyMMdd', $null).ToString("yyyy-MM-dd")
            #}
            #Else {$Date = $Null}
            
            $Output += New-Object PSObject -Property @{
                ComputerName    = $Env:ComputerName
                DisplayName     = $($thisSubKey.GetValue(“DisplayName”))
                DisplayVersion  = $($thisSubKey.GetValue(“DisplayVersion”))
                Publisher       = $($thisSubKey.GetValue(“Publisher”))
                InstallLocation = $($thisSubKey.GetValue(“InstallLocation”))
                InstallDate     = $($thisSubKey.GetValue(“InstallDate”))
                #InstallDate     = $Date
                UninstallString = $($thisSubKey.GetValue(“UninstallString”))
            }
        } 
        $Output | Where-Object {$_.DisplayName } | Sort-Object DisplayName | Select-Object ComputerName,DisplayName,DisplayVersion,Publisher,InstallLocation,InstallDate,UninstallString
 

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
                        Invoke-Command @Param_IC | Select-Object -Property * -ExcludeProperty PSComputerName,RunspaceID
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

