#requires -Version 3
Function Show-PKNestedProgressBars {
<#
.SYNOPSIS
    Demonstrates four nested progress bars displaying the hour, minute, second, and millisecond in a countdown to midnight

.DESCRIPTION
    Demonstrates four nested progress bars displaying the hour, minute, second, and millisecond in a countdown to midnight
    Returns no output
    
.NOTES        
    Name    : Function_Show-PKNestedProgressBars.ps1
    Author  : Paula Kingsley
    Created : 2017-11-29
    Version : 01.00.0000
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2017-11-29 - Created script based on markwragg's original

        TO DO Someday may work with custom end time parameter 
        
.LINK
    https://gist.github.com/markwragg/73addf16504caaf72da1633cdac57e68
#>
[Cmdletbinding()]
Param(

    #[Parameter(
    #    Mandatory = $False,
    #    HelpMessage = "End time (default is midnight)"
    #)]
    #[ValidateNotNullOrEmpty()]
    #[ValidateScript({
    #    If (($_ -as [datetime]) -and ($_ -as [datetime] -gt (Get-Date))) {$True}
    #    Else {$Host.UI.WriteErrorLine("Please select a date/time greater than now");$False}
    #})]
    #$EndTime# = (Get-Date).AddMinutes(5) 
    
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
    $ProgressPreference   = "Continue"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }


    $Desc = (Get-Date -hour 0 -Minute 0 -Second 0).ToLongDateString()
    $Stoptime = [datetime]::Today
    $Stop = ((Get-Date).hour -eq $Stoptime.hour)

    <#
    If (-not $EndTime) {
        #$Desc = "midnight"
        $Desc = (Get-Date -hour 0 -Minute 0 -Second 0).ToLongDateString()
        $Stoptime = [datetime]::Today
        $Stop = ((Get-Date).hour -eq $Stoptime.hour)
    }
    Else {
        $Desc = $EndTime
        $StopTime = Get-Date $EndTime
        #$Stop = ((Get-Date) -ge (Get-Date $EndTime))
        $Stop = ((Get-Date).hour -eq $Stoptime.hour)
    }
    #>
    
    # Console output
    $BGColor = $Host.UI.RawUI.BackgroundColor
    $Msg = "ACTION: Display nested progress bars counting down the hours, minutes, seconds, and milliseconds until $Stoptime"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}

Process{
    
    # I want to make this work with a custom end time but can't figure it out right now. Safekeeping here. 
    <#
    # Hours
    Do {
        $CurentHour = Get-Date
        Write-Progress -Activity "$($CurrentTime.hour) hours" -PercentComplete (((get-date).hour /23) * 100) -Status "$(24 - (get-date).hour) hours remaining"

        # Minutes
        Do {
            $CurentMinute = Get-Date
            Write-Progress -Id 1 -Activity "$($CurrentMinute.minute) minutes" -PercentComplete (((get-date).minute / 59) * 100) -Status "$(60 - (get-date).minute) minutes remaining" 
            
            # Seconds
            Do {
                $CurrentSec = Get-Date
                Write-Progress -Id 2 -Activity "$($CurrentSec.second) seconds" -PercentComplete (((get-date).second / 59) * 100) -Status "$(60 - (get-date).second) seconds remaining" 
                
                # Milliseconds
                Do {
                    $Second = (Get-Date).second
                    Write-Progress -Id 3 -Activity "$((get-date).millisecond) milliseconds" -Status "The time is $(get-date -f "HH:mm:ss")" `
                        -PercentComplete (((get-date).millisecond / 1000) * 100) -SecondsRemaining (86400 - (((get-date).hour * 60 * 60) + ((get-date).Minute * 60) + ((get-date).Second)))
                }
                Until ((get-date).second -ne $Second)
            }
            Until ((get-date).second -eq 0)
        }
        Until ((get-date).minute -eq 0)
    } 
    Until ((Get-Date).hour -eq $Stoptime.hour)

    #>

    
    # Hours
    Do {
        Write-Progress -Activity "$((get-date).hour) hours" -PercentComplete (((get-date).hour /23) * 100) -Status "$(24 - (get-date).hour) hours remaining to $Stoptime"

        # Minutes
        Do {
            Write-Progress -Id 1 -Activity "$((get-date).minute) minutes" -PercentComplete (((get-date).minute / 59) * 100) -Status "$(60 - (get-date).minute) minutes remaining" 
            
            # Seconds
            Do{
                Write-Progress -Id 2 -Activity "$((get-date).second) seconds" -PercentComplete (((get-date).second / 59) * 100) -Status "$(60 - (get-date).second) seconds remaining" 
                
                # Milliseconds
                Do{
                    $Second = (Get-Date).second
                    Write-Progress -Id 3 -Activity "$((get-date).millisecond) milliseconds" -Status "The time is $(get-date -f "HH:mm:ss")" `
                        -PercentComplete (((get-date).millisecond / 1000) * 100) -SecondsRemaining (86400 - (((get-date).hour * 60 * 60) + ((get-date).Minute * 60) + ((get-date).Second)))
                }
                Until ((get-date).second -ne $Second)
            }
            Until ((get-date).second -eq 0)
        }
        Until ((get-date).minute -eq 0)
    } 
    Until ((Get-Date) -eq $Stoptime)

}
} #end Show-PKNestedProgressbars

