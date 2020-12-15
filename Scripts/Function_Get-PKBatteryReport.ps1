#Requires -version 3
Function Get-PKBatteryReport {
<# 
.SYNOPSIS
    Runs powercfg to create an HTML report on the utilization and history of the local computer battery

.DESCRIPTION
    Runs powercfg to create an HTML report on the utilization and history of the local computer battery
    Returns an HTML file and optionally launches it in the default file handler

.NOTES        
    Name    : Function_Get-PKBatteryReport.ps1
    Created : 2020-08-04
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2020-08-04 - Created script

.LINK
    # https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/sophisticated-battery-report

.PARAMETER OutputFile
    Path to output file (default is current user's temp folder)

.PARAMETER Days
    Duration in days (default is 14)

.PARAMETER LaunchFile
    Launch HTML file after creation

.PARAMETER Quiet
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Get-PKBatteryReport -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                               
        ---           -----                               
        Verbose       True                                
        OutputPath    C:\Users\jbloggs\AppData\Local\Temp
        Days          14                                  
        LaunchFile    False                               
        Quiet         False                               
        ScriptName    Get-PKBatteryReport                 
        ScriptVersion 1.0.0                               


        [BEGIN: Get-PKBatteryReport] Create HTML report for 14-day battery history on LAPTOP14

        Battery life report saved to file path C:\Users\jbloggs\AppData\Local\Temp\LAPTOP14_2020-08-04_15-33.html.
        VERBOSE: Launching report using default HTML file handler

        [END Get-PKBatteryReport] Create HTML report for 14-day battery history on LAPTOP14


.EXAMPLE
    PS C:\> Get-PKBatteryReport -OutputPath 'G:\My Drive\reports' -Days 30 -Verbose -Quiet

        VERBOSE: PSBoundParameters: 
	
        Key           Value              
        ---           -----              
        OutputPath    G:\My Drive\reports
        Days          30                 
        Verbose       True               
        Quiet         True               
        LaunchFile    False              
        ScriptName    Get-PKBatteryReport
        ScriptVersion 1.0.0              



        VERBOSE: [BEGIN: Get-PKBatteryReport] Create HTML report for 30-day battery history on LAPTOP14
        VERBOSE: Battery life report saved to file path G:\My Drive\reports\LAPTOP14_2020-08-04_15-30.html.
        VERBOSE: HTML report file launch canceled
        VERBOSE: [END Get-PKBatteryReport] Create HTML report for 30-day battery history on LAPTOP14

        
#> 

[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param (
    [Parameter(
        Position = 0,
        HelpMessage = "Path to output file (default is current user's temp folder)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If (Test-Path -Path $_ -PathType Container -EA SilentlyContinue) {$True}})]
    [String] $OutputPath = $Env:Temp,

    [Parameter(
        HelpMessage = "Duration in days (default is 14)"
    )]
    [int] $Days = 14,

    [Parameter(
        HelpMessage = "Launch HTML file after creation"
    )]
    [switch]$LaunchFile,

    [Parameter(
        HelpMessage = "Hide all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet
)

Begin { 
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name

    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

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

    # Function to write an error as a string (no stacktrace), or an error, and options for prefix to string
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        If (-not $Quiet.IsPresent) {
            $Host.UI.WriteErrorLine("$Message")
        }
        Else {Write-Error "$Message"}
    }
    # Function to write a warning, with any error data, and options for prefix to string
    Function Write-MessageWarning {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Warning $Message
    }

    # Function to write a verbose message, collecting error data
    Function Write-MessageVerbose {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Verbose $Message
    }

    #endregion Functions

    #region Prerequisites

    Try {
        If (-not ($Null = Get-Command Powercfg.exe -ErrorAction SilentlyContinue)) {
            $Msg = "Powercfg.exe not found"
            Write-MessageError -Message $Msg -PrefixPrerequisites -PrefixError
            Break
        }
    }
    Catch {
        $Msg = "Unable to find Powercfg.exe"
        $Msg | Write-MessageError -PrefixPrerequisites -PrefixError
        Break            
    }

    # File name 
    $Filename = "$($Env:ComputerName)_$(Get-Date -f yyyy-MM-dd_HH-mm).html"
    $OutputFile = "$($OutputPath)\$($FileName)"

    #endregion  Prerequisites

    # Console output
    $Activity = "Create HTML report for $Days-day battery history on $Env:ComputerName"
    "[BEGIN: $ScriptName] $Activity" | Write-MessageInfo -FGColor Yellow -Title
    

} #end begin

Process {
    
    Write-Progress -Activity $Activity -CurrentOperation "Creating $OutputFile"

    Try {
        $RunReport = powercfg /batteryreport /output $OutputFile /duration $Days

        If ($Null = Get-Item -Path $OutputFile -ErrorAction SilentlyContinue) {
            $Msg = $RunReport
            $Msg | Write-MessageInfo -FGColor Green

            $ConfirmMsg = "`n`n`tLaunch output file?`n`n"
            If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
                Try {
                    $Msg = "Launching report using default HTML file handler"
                    $Msg | Write-MessageVerbose
                    Start-Process -FilePath $OutputFile -ErrorAction Stop

                }
                Catch {
                    $Msg = "Failed to launch output file"
                    $Msg | Write-MessageError -PrefixError
                }
            }
            Else {
                $Msg = "HTML report file launch canceled"
                $Msg | Write-MessageVerbose 
            }
        }
    }
    Catch {
        $Msg = "Failed to run powercfg"
        $Msg | Write-MessageError -PrefixError
    }


}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed
    "[END $ScriptName] $Activity" | Write-MessageInfo -FGColor Yellow -Title
}

} # end Get-PKBatteryReport

