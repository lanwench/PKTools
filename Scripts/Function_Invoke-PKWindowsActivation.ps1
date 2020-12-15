#requires -Version 3
Function Invoke-PKWin2019Activation {
<#
.SYNOPSIS 
    Changes the product key and activates Windows 2019 Standard or Datacenter

.DESCRIPTION
    Changes the product key and activates Windows 2019 Standard or Datacenter
    First detects current activation status
    Accepts Standard and Datacenter product keys 
    Uses Invoke-Command and runs as a remote PSJob
    Accepts pipeline input

.NOTES
    Name      : Function_Invoke-PKWin2019Activation.ps1
    Created   : 2020-11-12
    Author    : Paula Kingsley
    Version   : 01.00.0000
    Changelog :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2020-11-12 - Created script

.EXAMPLE
    PS C:\> Get-ADComputer -SearchBase "OU=Servers,DC=domain,DC=local" -SearchScope Subtree -filter * -properties DNSHostname).DNSHostname | 
        Invoke-Win2019Activation -StandardKey "ABC12-DEF34-GHI56-KLM78-NOP90" -DatacenterKey "ABC09-DEF87-GHI65-KLM43-NOP21" -Credential (Get-Credential domain\adminuser) -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                                    
        ---           -----                                    
        StandardKey   ABC12-DEF34-GHI56-KLM78-NOP90            
        DatacenterKey ABC09-DEF87-GHI65-KLM43-NOP21            
        Credential    System.Management.Automation.PSCredential
        Verbose       True                                     
        ComputerName                                           
        PipelineInput True                                     
        ScriptName    Invoke-Win2019Activation                 
        ScriptVersion 1.0.0                                    


        VERBOSE: [sql1.domain.local] Test WinRM
        VERBOSE: [sql1.domain.local] Invoke scriptblock as job 'Key_sql1.domain.local'
        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        31     Key_sql.doma... RemoteJob       Running       True            sql1.nl... ...                      
        
        VERBOSE: [webprod2.domain.local] Test WinRM
        VERBOSE: [webprod2.domain.local] Invoke scriptblock as job 'Key_webprod2.domain.local'
        33     Key_webprod... RemoteJob       Running       True            webprod2.dom... ...                      
        
        VERBOSE: [sandbox.domain.local] Test WinRM
        VERBOSE: [sandbox.domain.local] Invoke scriptblock as job 'Key_sandbox.domain.local'
        35     Key_sandbox... RemoteJob       Running       True            sandbox.do... ...                      
        
        VERBOSE: [vcenter1.domain.local] Test WinRM
        VERBOSE: [vcenter1.domain.local] Invoke scriptblock as job 'Key_vcenter1.domain.local'
        37     Key_vcenter... RemoteJob       Running       True            vcenter1.do... ...                      
        
        
        PS C:> Get-Job 

        Id     Name        PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----        -------------   -----         -----------     --------             -------                  
        31     Key_sql1... RemoteJob       Completed     True            sql1.domain.local... ...                      
        33     Key_webp... RemoteJob       Completed     True            webprod1.domain.l... ...                      
        35     Key_sand... RemoteJob       Running       True            sandbox.domain.lo... ...                      
        37     Key_vcen... RemoteJob       Running       True            vcenter1.domain.l... ...                      
        

        PS C:\> Get-Job 31,33,35,37,39 | Receive-Job | Select -Property * -ExcludeProperty PSComputerName,RunspaceID


        Computername   : SQL1
        IsActivated    : True
        ProductKey     : ABC12-DEF34-GHI56-KLM78-NOP90
        WindowsVersion : Microsoft Windows Server 2019 Standard
        ApplicationID  : 55c92734-d682-4d71-983e-d6ec3f16059f
        LicenseStatus  : Licensed
        Messages       : Successfully applied product key and activated Windows
        
        Computername   : WEBPROD2
        IsActivated    : False
        ProductKey     : -
        WindowsVersion : Microsoft Windows Server 2019 Datacenter
        ApplicationID  : 55c92734-d682-4d71-983e-d6ec3f16059f
        LicenseStatus  : Notification
        Messages       : Successfully applied product key; activation failed
        
        Computername   : SANDBOX
        IsActivated    : False
        ProductKey     : -
        WindowsVersion : Microsoft Windows Server 2016 Standard
        ApplicationID  : 55c92734-d682-4d71-983e-d6ec3f16059f
        LicenseStatus  : Notification
        Messages       : Script requires Windows Server 2019 Standard or Datacenter
        
        Computername   : VCENTER1
        IsActivated    : True
        ProductKey     : -
        WindowsVersion : Microsoft Windows Server 2019 Datacenter
        ApplicationID  : 55c92734-d682-4d71-983e-d6ec3f16059f
        LicenseStatus  : Licensed
        Messages       : Windows is already activated
        
        

#>
[Cmdletbinding()]
Param(
    [Parameter(
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        Mandatory = $True
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name","DNSHostName")]
    [string]$ComputerName,

    [Parameter(
        Mandatory = $True
    )]
    [ValidateNotNullOrEmpty()]
    [string]$StandardKey = "TDH83-8NKV9-76PK3-JGJMM-F9FBQ",

    [Parameter(
        Mandatory = $True
    )]
    [ValidateNotNullOrEmpty()]
    [string]$DatacenterKey = "NKV7F-H3644-3V9W4-DRTKC-HH8V7",

    [Parameter(
        Mandatory = $True
    )]
    [ValidateNotNullOrEmpty()]
    [PSCredential]$Credential

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $ScriptBlock = {
        Param($StandardKey,$DatacenterKey)
        
        $Output = [pscustomobject]@{
            Computername     = $env:COMPUTERNAME
            IsActivated      = $Null
            ProductKey       = $Null
            WindowsVersion   = $Null
            ApplicationID    = $Null
            LicenseStatus    = $Null
            Messages         = $Null
        }

        $LicenseTable = @{
            0 = "Unlicensed"
            1 = "Licensed"
            2 = "OOB Grace"
            3 = "OOT Grace"
            4 = "Non-Genuine Grace"
            5 = "Notification"
            6 = "Extended Grace"
        }

        # Get the Windows version via WMI
        Try {

            $OSVer = (Get-WMIObject -Query "Select Caption FROM Win32_Operatingsystem" -ErrorAction Stop).Caption
            $Output.WindowsVersion = $OSVer

            # Get the activation status via WMI
            Try {
                $Status = Get-WmiObject SoftwareLicensingProduct -ErrorAction Stop | Where-Object {$_.PartialProductKey} 
                $Output.ApplicationID = $Status.ApplicationId
                $Output.LicenseStatus = $LicenseTable[$Status.LicenseStatus -as [int]]
                    
                If ($Status.LicenseStatus -eq 1) {
                    $Output.ProductKey = "-"
                    $Output.IsActivated = $True
                    $Output.Messages = "Windows is already activated"
                }
                Else {
                    $Output.IsActivated = $False
                    Switch -regex ($OSVer) {
                        "2019 Datacenter" {$Key = $Using:DatacenterKey}
                        "2019 Standard" {$Key = $Using:StandardKey}
                        Default {}
                    }
                    
                    If ($Key) {
                        $Output.ProductKey = $Key

                        # Set the product key
                        Try {
                            $Service = Get-WMIObject -Query "select * from SoftwareLicensingService" -ErrorAction Stop
                            $Set = $Service.InstallProductKey($Key)
                            $Refresh = $Service.RefreshLicenseStatus()
                            
                            Start-Sleep 3
                            $Status = Get-WmiObject SoftwareLicensingProduct -ErrorAction Stop | Where-Object {$_.PartialProductKey} 
                            $Output.LicenseStatus = $LicenseTable[$Status.LicenseStatus -as [int]]

                            If ($Status.LicenseStatus -eq 1) {
                                $Output.IsActivated = $True
                                $Output.Messages = "Successfully applied product key and activated Windows"
                            }
                            Else {
                                $Output.Messages = "Successfully applied product key; activation failed"
                            }
                        }
                        Catch {
                            $Output.Messages = $_.Exception.Message
                        }
                    }
                    Else {
                        $Output.Messages = "Script requires windows Server 2019 Datacenter or Standard"
                    }
                }
            }
            Catch {
                $Output.Messages = $_.Exception.Message
            }
        }
        Catch {
            $Output.Messages = $_.Exception.Message
        }

        Write-Output $Output
    }
    
    # Splats

    $Param_TestWSMan = @{}
    $Param_TestWSMAN = @{
        ComputerName   = $Null
        Authentication = "Kerberos"
        Credential     = $Credential
        ErrorAction    = "Stop"
    }

    $Param_InvokeCommand = @{}
    $Param_InvokeCommand = @{
        ComputerName   = $Null
        Credential     = $Credential
        Authentication = "Kerberos"
        Scriptblock    = $ScriptBlock
        ArgumentList   = $StandardKey,$DatacenterKey
        AsJob          = $True
        JobName        = $Null
        ErrorAction    = "Stop"
    }

}
Process {

    Foreach ($Computer in $ComputerName) {
        
        $Computer = $Computer.Trim()
        $Param_TestWSMan.ComputerName = $Param_InvokeCommand.ComputerName = $Computer
        $Param_InvokeCommand.JobName = "Key_$Computer"
        
        Write-Verbose "[$Computer] Test WinRM"
        Try {
            
            $Null = Test-WSMan @Param_TestWSMAN

            Write-Verbose "[$Computer] Invoke scriptblock as job 'Key_$Computer'"
            Try {
                Invoke-Command @Param_InvokeCommand
            }
            Catch {
                Write-Error "[$Computer] Failed to invoke remote scriptblock ($($_.Exception.Message))"
            }
        }
        Catch {
            Write-Error "[$Computer] Failed to connect using WinRM and Kerberos authentication ($($_.Exception.Message))"
        }
    }

}
End {

}
}