#requires -Version 3
Function Get-PKWindowsActivationStatus {
<#
.SYNOPSIS 
    Gets the activation status for one or more remote Windows computers

.DESCRIPTION
    Gets the activation status for one or more remote Windows computers
    Uses Invoke-Command and runs as a remote PSJob
    Accepts pipeline input

.NOTES
    Name      : Function_Get-PKWindowsActivationStatus.ps1
    Created   : 2020-11-12
    Author    : Paula Kingsley
    Version   : 01.00.0000
    Changelog :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2020-11-12 - Created script

.EXAMPLE
    PS C:\> Get-ADComputer -SearchBase "OU=Servers,DC=domain,DC=local" -SearchScope Subtree -filter * -properties DNSHostname).DNSHostname | 
        
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
        
        $Output = [pscustomobject]@{
            Computername     = $env:COMPUTERNAME
            IsActivated      = $Null
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
                    $Output.IsActivated = $True
                    $Output.Messages = "Windows is already activated"
                }
                Else {
                    $Output.IsActivated = $False
                    $Output.Messages = "Windows is not activated"
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