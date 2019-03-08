#Requires -version 3
Function Get-PKWindowsLicenseInfo {
<# 
.SYNOPSIS
    Uses good old slmgr.vbs & WMI to retrieve Windows licensing on a local or remote computer, interactively or as a PSJob

.DESCRIPTION
    Uses good old slmgr.vbs & WMI to retrieve Windows licensing on a local or remote computer, interactively or as a PSJob
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Get-PKWindowsLicenseInfo.ps1
    Created : 2019-02-26
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-02-26 - Created script

.LINK
    https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/managing-windows-license-key

.LINK
    https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/managing-windows-license-key-part-2

.PARAMETER ComputerName
    One or more computer names

.PARAMETER Credential
    Valid credentials on target (default is current user credentials)

.PARAMETER AsJob
    Invoke command as a PSjob

.PARAMETER JobPrefix
    Prefix for job name (default is 'License')

.PARAMETER ConnectionTest
    Run WinRM or ping connectivity test prior to Invoke-Command, or no test (default is None)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKWindowsLicenseInfo -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Verbose               True                                     
        ComputerName          {LAPTOP-212}                        
        Credential            System.Management.Automation.PSCredential
        AsJob                 False                                    
        JobPrefix             License                                  
        ConnectionTest        None                                     
        SuppressConsoleOutput False                                    
        ScriptName            Get-PKWindowsLicenseInfo                 
        ScriptVersion         1.0.0                                    

        START  : Invoke scriptblock to return Windows licensing details
        VERBOSE: [LAPTOP-212] Invoke command

        ComputerName                  : LAPTOP-212
        Name                          : Windows(R), Enterprise edition
        Description                   : Windows(R) Operating System, VOLUME_MAK channel
        Activation ID                 : 2ffd8952********************************
        Application ID                : 55c92734********************************
        Extended PID                  : 03612-03********************************
        Product Key Channel           : Volume
        Installation ID               : 31404061********************************
        Use License URL               : https
        Validation URL                : https
        Partial Product Key           : RVV2F
        License Status                : Licensed
        Remaining Windows rearm count : 1001
        Remaining SKU rearm count     : 1001
        Trusted time                  : 2019-02-26 10
        WMI Operating System          : Microsoft Windows 10 Enterprise
        WMI Operating System Version  : 10.0.17134.471
        WMI Product Key               : V****-9****-6****-B****-P****

.EXAMPLE
    PS C:\> $Arr | Get-PKWindowsLicenseInfo -ConnectionTest WinRM -Credential $Credential -AsJob
        
        START  : Invoke scriptblock to return Windows licensing details as PSJob

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        43     License_ops-... RemoteJob       Running       True            ops-pkbastion-1.d... ...                      
        45     License_ops-... RemoteJob       Running       True            ops-wsus-201.doma... ...                      
        [trackid-metaprod-2.domain.local] WinRM connection test failure
        END    : 2 job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output
                
        PS C:\> Get-Job lic* | Receive-Job -Keep

        ComputerName                  : OPS-PKBASTION-1
        Name                          : Windows(R), ServerDatacenter edition
        Description                   : Windows(R) Operating System, VOLUME_MAK channel
        Activation ID                 : 641f81b2-6******************************
        Application ID                : 55c92734-d******************************
        Extended PID                  : 06401-0253******************************
        Product Key Channel           : Volume
        Installation ID               : 2226461831******************************
        Use License URL               : https
        Validation URL                : https
        Partial Product Key           : T29CW
        License Status                : Licensed
        Remaining Windows rearm count : 999
        Remaining SKU rearm count     : 1000
        Trusted time                  : 2/26/2019 10
        WMI Operating System          : Microsoft Windows Server 2012 R2 Datacenter
        WMI Operating System Version  : 6.3.9600.19101
        WMI Product Key               : (not available)
        PSComputerName                : ops-pkbastion-1.domain.local
        RunspaceId                    : 28128814-f448-435d-9d32-3788737524e5

        ComputerName                  : OPS-WSUS-201
        Name                          : Windows(R), ServerDatacenter edition
        Description                   : Windows(R) Operating System, VOLUME_MAK channel
        Activation ID                 : 641f81b2-6******************************
        Application ID                : 55c92734-d******************************
        Extended PID                  : 06401-0253******************************
        Product Key Channel           : Volume
        Installation ID               : 0169835317******************************
        Use License URL               : https
        Validation URL                : https
        Partial Product Key           : 27HJ8
        License Status                : Licensed
        Remaining Windows rearm count : 999
        Remaining SKU rearm count     : 1000
        Trusted time                  : 2/26/2019 1
        WMI Operating System          : Microsoft Windows Server 2012 R2 Datacenter
        WMI Operating System Version  : 6.3.9600.19101
        WMI Product Key               : (not available)
        PSComputerName                : ops-wsus-201.domain.local
        RunspaceId                    : a8f74c69-4b0a-467f-a5e4-fac3d8305e14      
        
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
        #Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName = $env:COMPUTERNAME,

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
        HelpMessage = "Prefix for job name (default is 'License')"
    )]
    [String] $JobPrefix = "License",

    [Parameter(
        HelpMessage="Test to run prior to Invoke-Command: WinRM, Ping, or None (default None)"
    )]
    [ValidateSet("WinRM","Ping","None")]
    [string] $ConnectionTest = "None",

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
    
    #region Scriptblock for Invoke-Command

    $ScriptBlock = {
        Try {

            # Find slmgr, create the command line, and execute it
            $Path = Get-Command slmgr.vbs | Select-Object -ExpandProperty Source
            $Data = cscript.exe //Nologo $Path /dlv
            
            # Split the results at the colon symbol & create an array object
            $OutputObj = $Data | ConvertFrom-Csv -Delimiter ":"  -Header "Key","Value"
            
            # Try to get WMI product key & OS version
            $WMILicense = (Get-WmiObject -Class SoftwareLicensingService | Select-Object @{N="ProductKey";E={$_.OA3xOriginalProductKey}},@{N="OS";E={(Get-WMIObject -Class Win32_OperatingSystem).Caption}},Version)
            If (-not $WMILicense.ProductKey) {$WMILicense.ProductKey = "(not available)"}

            # Selection for output order
            $Select = @("ComputerName,Name,Description,Activation ID,Application ID,Extended PID,Product Key Channel,Installation ID,Use License URL,Validation URL,Partial Product Key,License Status,Remaining Windows rearm count,Remaining SKU rearm count,Trusted time,WMI Operating System,WMI Operating System Version,WMI Product Key") -split(",")

            # Create a new hashtable
            $OutputHash = @{}
            $OutputObj | Foreach-Object {
                $Curr = $_
                $OutputHash.Add($_.Key,$_.Value)
            }

            # Create a PSObject from the hashtable, adding the two additional properties, then output with the selection order
            New-Object PSObject -Property $OutputHash | 
                Add-Member -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME -Force -PassThru | 
                    Add-Member -MemberType NoteProperty -Name "WMI Product Key" -Value $WMILicense.ProductKey -Force -PassThru | 
                        Add-Member -MemberType NoteProperty -Name "WMI Operating System" -Value $WMILicense.OS -Force -PassThru | 
                            Add-Member -MemberType NoteProperty -Name "WMI Operating System Version" -Value $WMILicense.Version -Force -PassThru | 
                            Select-Object $Select
        }
        Catch {
            Throw $_.Exception.Message
        }
    } #end scriptblock

    #endregion Scriptblock for Invoke-Command

    #region Functions

    Function Test-WinRM{
        Param($Computer,$Credential)
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
    $Activity = "Invoke scriptblock to return Windows licensing details"
    If ($AsJob.IsPresent) {
        $Activity = "$Activity as PSJob"
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
    $Msg = "START  : $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
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
                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                        If ($Null = Test-Ping -Computer $Computer) {
                            
                            $Msg = "Successfully pinged computer"
                            Write-Verbose "[$Computer] $Msg"

                            $Continue = $True
                        }
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
                    $ConfirmMsg = "`n`n`t$Msg`n`n"
                    If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                        If ($Null = Test-WinRM -Computer $Computer -Credential $Credential) {
                            
                            $Msg = "WinRM connection test successful"
                            Write-Verbose "[$Computer] $Msg"

                            $Continue = $True
                        }
                        Else {
                            $Msg = "WinRM connection test failure"
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
                        $Job
                    }
                    Else {
                        $Results = Invoke-Command @Param_IC                                              
                        Write-Output $Results | Select -Property * -ExcludeProperty PSComputerName,RunspaceID
                    }
                }
                Catch {
                    $Msg = "Operation failed"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR : [$Computer] $Msg")
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                Write-Warning "[$Computer] $Msg"
            }
        
        } #end if proceeding with script
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($AsJob.IsPresent) {

        If ($Jobs.Count -gt 0) {
            $Msg = "$($Jobs.Count) job(s) created; run 'Get-Job -Id # | Wait-Job | Receive-Job' to view output"
            If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"END    : $Msg")}
            Else {Write-Verbose "END    : $Msg"}
        }
        Else {
            $Msg = "No jobs created"
            Write-Warning "$Msg"
        }
    } #end if AsJob
}

} # end Get-PKWindowsProduct Key

