Function Repair-PKWindowsDISM {

[CmdletBinding(
     DefaultParameterSetName = "ByType",
     SupportsShouldProcess = $True,
     ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Name of computer (separate multiple computers with commas)"
    )]
    [Alias("FQDN","HostName","Computer","VM")]
    [String[]] $ComputerName,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Valid credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty

)

Begin {
    
    $Scriptblock = {
        $Output = New-Object PSObject -Property @{
            ComputerName = $env:COMPUTERNAME
            Repaired = "Error"
            Messages = "Error"
        }
        $Cmd = "DISM.exe /Online /Cleanup-image /Restorehealth"
        Try {
            
            $Results = Invoke-Expression -Command $Cmd -ErrorAction Stop 
            If ($Results -match "The restore operation completed successfully") {
                $Output.Repaired = $True
            }
            Else {
                $Output.Repaired = $False
            }
            $Output.Messages = $Results
        }
        Catch {
            $Output.Messages = $_.Exception.Message
        }
        Write-Output $Output | Select Computername,Repaired,Messages
    }

}
Process {
    
}

}


break

# https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/use-dism-in-windows-powershell-s14

$Repair = Repair-WindowsImage -Online -RestoreHealth -NoRestart -LogLevel WarningsInfo -LogPath C:\Users\PKINGS~1\AppData\Local\Temp\CE-AGILEDEV-1_RestoreHealth_2017-09-22_12-13-20.log -Verbose


pause

[switch]$Continue = $False

Try {
    
    [version]$MinVer = "6.3" 
    If (([version](Get-WMIObject -Class win32_operatingsystem -ErrorAction Stop).version) -lt $MinVer) {
        $Continue = $False
    }
    Else {$Continue = $True}
   
}
Catch {

}





# http://woshub.com/how-to-repair-the-component-store-in-windows-8/



    $InitialValue = "Error"
    $Output = New-Object PSObject -Property ([ordered]@{
        ComputerName            = $Env:ComputerName
        RanHealthScan           = $InitialValue
        RanHealthCheck          = $InitialValue
        RanHealthRestore        = $InitialValue
        HealthScan              = $InitialValue
        HealthCheck             = $InitialValue
        HealthRestore           = $InitialValue
        ResultsHealthScan       = $InitialValue
        ResultsHealthCheck      = $InitialValue
        ResultsHealthRestore    = $InitialValue
        Logs                    = $InitialValue
        Messages                = $InitialValue
    })

    Try {
        $M = Get-Module DISM -ListAvailable -ErrorAction Stop
        $Continue = $True
    }
    Catch {
        $Msg = "Module 'DISM' not found"
        $Output.Messages = $Msg

        $Continue = $False
        Write-Output $Output
        Break
    }


If ($Continue.IsPresent) {

    Try {
        $VerbosePreference = "Continue"
        $ProgressPreference = "Continue"

        $Date = (Get-Date -f yyyy-MM-dd_hh-mm-ss)
    
        $Logfile_Scan = "$Env:Temp\$Env:ComputerName`_ScanHealth_$Date.log"
        $Logfile_Check = "$Env:Temp\$Env:ComputerName`_CheckHealth_$Date.log"
        $Logfile_Restore = "$Env:Temp\$Env:ComputerName`_RestoreHealth_$Date.log"
    
        $Logs = @()

        $Total = 3
        $Current = 0
        $Param_WP = @{}
        $Param_WP = @{
            CurrentOperation = $Null
            Activity         = "Scan and repair Windows Image"
            Status           = "Working"
            PercentComplete  = $Null
        }
        $StartTime = Get-Date

        $Msg = "Scan health (log $Logfile_Scan)"
        Write-Verbose $Msg
        $Current ++

        $Param_WP.CurrentOperation = $Msg
        $Param_WP.PercentComplete = ($Current / $Total * 100)
        Write-Progress @Param_WP
    
        $ScanHealth = Repair-WindowsImage –Online –ScanHealth –LogPath $Logfile_Scan -LogLevel Errors -Verbose -ErrorAction Stop
        $Output.RanHealthScan = $True
        $Output.HealthScan = $ScanHealth.ImageHealthState
        $Output.ResultsHealthScan = ($ScanHealth | Select ImageHealthState,Path,Online,WinPath,SysDrivePath,RestartNeeded,LogLevel,LogPath) -join(", ")
        $Logs += $Logfile_Scan
    
        
        $Msg = "Check health (log $Logfile_Check)"
        Write-Verbose $Msg
        $Current ++

        $Param_WP.CurrentOperation = $Msg
        $Param_WP.PercentComplete = ($Current / $Total * 100)
        Write-Progress @Param_WP
    
        $CheckHealth = Repair-WindowsImage –Online –LogPath $LogFile_Check -LogLevel Errors -CheckHealth -Verbose -ErrorAction Stop
        $Output.RanHealthCheck = $True
        $Output.HealthCheck = $CheckHealth.ImageHealthState
        $Output.ResultsHealthCheck = ($CheckHealth | Select ImageHealthState,Path,Online,WinPath,SysDrivePath,RestartNeeded,LogLevel,LogPath) -join(", ")
        $Logs += $Logfile_Check

        
        $Msg = "Repair health (log $Logfile_Restore)"
        Write-Verbose $Msg
        $Current ++
        $Param_WP.CurrentOperation = $Msg
        $Param_WP.PercentComplete = ($Current / $Total * 100)
        Write-Progress @Param_WP
    
        $RestoreHealth = Repair-WindowsImage –Online –LogPath $LogFile_Restore -LogLevel WarningsInfo -RestoreHealth -Verbose -ErrorAction Stop
        $Output.RanHealthRestore = $True
        $Output.HealthRestore = $RestoreHealth.ImageHealthState
        $Output.ResultsHealthRestore = ($RestoreHealth | Select ImageHealthState,Path,Online,WinPath,SysDrivePath,RestartNeeded,LogLevel,LogPath) -join(", ")
        $Logs += $Logfile_Restore

        Write-Progress -Activity $Param_WP.Activity -Complete

        $EndTime = Get-Date
        $Elapsed = ($EndTime - $StartTime)
        $Msg = "Operation completed in $($Elapsed.Minutes)m, $($Elapsed.Seconds)s"
        Write-Output $Msg 

        $Output.Logs = $Logs
        $Output.Messages = $Msg
    }
    Catch {

        $EndTime = Get-Date
        $Elapsed = ($EndTime - $StartTime)
        $Msg = "Operation failed after $($Elapsed.Minutes)m, $($Elapsed.Seconds)s"    
        $ErrorDetails =  $_.Exception.Message
        $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")

        $Output.Logs = $Logs
        $Output.Messages = "$Msg`n$ErrorDetails"
    }

    Write-Output $Output
}

#Repair-WindowsImage –Online -RestoreHealth -Source C:OfflineMountWindowsWinSxS -LimitAccess -LogPath C:TempRepair.log 